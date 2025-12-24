# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System.Collections.Generic import List

import System.Windows.Forms as WF
from System.Windows.Forms import (
    Form,
    Button,
    TextBox,
    Label,
    AnchorStyles,
    FormStartPosition,
    ComboBox,
    ComboBoxStyle,
)
from System.Drawing import Size, Point, Font, FontStyle


__persistentengine__ = True


class ScheduleActionHandler(IExternalEventHandler):
    def __init__(self, form):
        self.form = form
        self._pending_action = None

    @property
    def has_pending_action(self):
        return self._pending_action is not None

    def set_action(self, action_name):
        self._pending_action = action_name

    def Execute(self, uiapp):
        action_name = self._pending_action
        self._pending_action = None

        if not action_name:
            return

        try:
            if action_name == "search":
                self.form.run_search()
            elif action_name == "select":
                self.form.run_select()
            elif action_name == "close":
                self.form.perform_close_from_handler()
        except Exception as exc:
            try:
                WF.MessageBox.Show(
                    u"Ошибка при выполнении действия: {0}".format(exc),
                    u"Поиск по спецификации",
                    WF.MessageBoxButtons.OK,
                    WF.MessageBoxIcon.Error,
                )
            except Exception:
                pass

    def GetName(self):
        return "SearchInScheduleExternalHandler"


search_form_instance = None


class SearchInScheduleForm(Form):
    """Поиск в спецификации и отметка найденных элементов."""

    FOUND_PARAM_NAME = u"ADSK_Позиция"
    FOUND_PARAM_VALUE = u"Найдено"
    TRANSACTION_TITLE = u"Отметить найденные элементы"
    RESTORE_TITLE = u"Восстановить ADSK_Позиция"

    def __init__(self, doc, uidoc):
        Form.__init__(self)

        self.doc = doc
        self.uidoc = uidoc
        self.DB = DB

        self.field_infos = []  # данные по полям спецификации
        self.filter_field_infos = []  # список для фильтрации по полям
        self.system_param_names = []  # доступные системные параметры
        self.current_system_param = None
        self.original_positions = {}  # ElementId -> исходное значение ADSK_Позиция
        self._is_closing_via_handler = False

        self.Text = u"Поиск в спецификации"
        self.StartPosition = FormStartPosition.CenterScreen
        self.Size = Size(750, 160)
        self.MinimumSize = Size(900, 150)
        self.TopMost = False
        self.ShowInTaskbar = False

        self.mono_font = Font("Consolas", 9.0, FontStyle.Regular)

        # Поле поиска
        self.lblSearch = Label()
        self.lblSearch.Text = u"Искать текст:"
        self.lblSearch.AutoSize = True
        self.lblSearch.Location = Point(10, 15)

        self.tbSearch = TextBox()
        self.tbSearch.Location = Point(110, 12)
        self.tbSearch.Width = 380
        self.tbSearch.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        self.lblField = Label()
        self.lblField.Text = u"Поле:"
        self.lblField.AutoSize = True
        self.lblField.Location = Point(510, 15)
        self.lblField.Anchor = AnchorStyles.Top | AnchorStyles.Right

        self.cbField = ComboBox()
        self.cbField.DropDownStyle = ComboBoxStyle.DropDownList
        self.cbField.Location = Point(560, 12)
        self.cbField.Width = 190
        self.cbField.Anchor = AnchorStyles.Top | AnchorStyles.Right

        self.btnSearch = Button()
        self.btnSearch.Text = u"Найти"
        self.btnSearch.Location = Point(760, 10)
        self.btnSearch.Width = 90
        self.btnSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right

        # Системные фильтры
        self.lblSystemParam = Label()
        self.lblSystemParam.Text = u"Системный параметр:"
        self.lblSystemParam.AutoSize = True
        self.lblSystemParam.Location = Point(10, 45)

        self.cbSystemParam = ComboBox()
        self.cbSystemParam.DropDownStyle = ComboBoxStyle.DropDownList
        self.cbSystemParam.Location = Point(150, 42)
        self.cbSystemParam.Width = 200
        self.cbSystemParam.Anchor = AnchorStyles.Top | AnchorStyles.Left

        self.lblSystemValue = Label()
        self.lblSystemValue.Text = u"Значение:"
        self.lblSystemValue.AutoSize = True
        self.lblSystemValue.Location = Point(370, 45)

        self.cbSystemValue = ComboBox()
        self.cbSystemValue.DropDownStyle = ComboBoxStyle.DropDownList
        self.cbSystemValue.Location = Point(440, 42)
        self.cbSystemValue.Width = 200
        self.cbSystemValue.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        # Кнопки действий
        self.btnSelect = Button()
        self.btnSelect.Text = u"Выбрать в модели"
        self.btnSelect.Width = 170
        self.btnSelect.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.btnClose = Button()
        self.btnClose.Text = u"Закрыть"
        self.btnClose.Width = 100
        self.btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self._update_bottom_buttons_positions()

        # События
        self.btnSearch.Click += self.on_search_click
        self.btnSelect.Click += self.on_select_click
        self.btnClose.Click += self.on_close_click
        self.cbSystemParam.SelectedIndexChanged += self.on_system_param_changed
        self.FormClosing += self.on_form_closing

        # Добавление контролов
        self.Controls.Add(self.lblSearch)
        self.Controls.Add(self.tbSearch)
        self.Controls.Add(self.lblField)
        self.Controls.Add(self.cbField)
        self.Controls.Add(self.btnSearch)

        self.Controls.Add(self.lblSystemParam)
        self.Controls.Add(self.cbSystemParam)
        self.Controls.Add(self.lblSystemValue)
        self.Controls.Add(self.cbSystemValue)

        self.Controls.Add(self.btnSelect)
        self.Controls.Add(self.btnClose)

        # Инициализация данных из активной спецификации
        view = self.doc.ActiveView
        if isinstance(view, self.DB.ViewSchedule):
            self.build_field_info(view)
            self.map_body_columns(view)
            self.build_filter_list()
            self.build_system_param_list(view)

        # ExternalEvent для запуска действий из UI-потока Revit
        self.ext_handler = ScheduleActionHandler(self)
        self.ext_event = ExternalEvent.Create(self.ext_handler)

    def _update_bottom_buttons_positions(self):
        margin_bottom = 35
        margin_right = 10
        bottom_y = self.ClientSize.Height - margin_bottom

        self.btnClose.Location = Point(
            self.ClientSize.Width - self.btnClose.Width - margin_right, bottom_y
        )
        self.btnSelect.Location = Point(
            self.btnClose.Left - self.btnSelect.Width - 5, bottom_y
        )

    def OnResize(self, args):
        Form.OnResize(self, args)
        try:
            self._update_bottom_buttons_positions()
        except Exception:
            pass

    # ---------- Поля спецификации ----------

    def build_field_info(self, view_schedule):
        self.field_infos = []

        definition = view_schedule.Definition
        doc = self.doc
        field_count = definition.GetFieldCount()

        for i in range(field_count):
            field = definition.GetField(i)

            is_hidden = False
            try:
                is_hidden = field.IsHidden
            except Exception:
                is_hidden = False

            col_index = i
            try:
                tmp = field.ColumnIndex
                if tmp >= 0:
                    col_index = tmp
            except Exception:
                pass

            title = u""
            try:
                sched_field = field.GetSchedulableField()
                if sched_field:
                    pname = sched_field.GetName(doc)
                    if pname:
                        title = pname
            except Exception:
                title = u""

            if not title:
                try:
                    title = field.ColumnHeading or u""
                except Exception:
                    title = u""

            if not title:
                title = u"Поле {0}".format(i + 1)

            self.field_infos.append(
                {
                    "col_index": col_index,
                    "title": title,
                    "is_hidden": is_hidden,
                    "body_col": None,
                }
            )

        self.field_infos.sort(key=lambda x: x["col_index"])

    def map_body_columns(self, view_schedule):
        table_data = view_schedule.GetTableData()
        body = table_data.GetSectionData(self.DB.SectionType.Body)
        first_body_col = body.FirstColumnNumber

        visible_infos = [fi for fi in self.field_infos if not fi["is_hidden"]]
        for idx, info in enumerate(visible_infos):
            info["body_col"] = first_body_col + idx

    def build_filter_list(self):
        self.cbField.Items.Clear()
        self.filter_field_infos = []

        self.cbField.Items.Add(u"По всем полям")
        self.filter_field_infos.append(None)

        for info in self.field_infos:
            title = info["title"]
            display_title = (
                u"{0} (скрыто)".format(title) if info["is_hidden"] else title
            )
            self.cbField.Items.Add(display_title)
            self.filter_field_infos.append(info)

        if self.cbField.Items.Count > 0:
            self.cbField.SelectedIndex = 0

    def get_selected_filter_info(self):
        idx = self.cbField.SelectedIndex
        if idx < 0 or idx >= len(self.filter_field_infos):
            return None
        return self.filter_field_infos[idx]

    # ---------- Системные параметры ----------

    def build_system_param_list(self, view_schedule):
        candidate_names = [u"Секция_Имя", u"ADSK_Система_Имя", u"Система"]

        self.system_param_names = []

        collector = (
            self.DB.FilteredElementCollector(self.doc, view_schedule.Id)
            .WhereElementIsNotElementType()
        )

        for name in candidate_names:
            has_value = False
            for elem in collector:
                try:
                    p = elem.LookupParameter(name)
                    if p and p.StorageType == self.DB.StorageType.String:
                        val = p.AsString()
                        if val:
                            has_value = True
                            break
                except Exception:
                    continue
            if has_value:
                self.system_param_names.append(name)

        self.cbSystemParam.Items.Clear()
        self.cbSystemValue.Items.Clear()

        if not self.system_param_names:
            self.cbSystemParam.Items.Add(u"(нет доступных параметров)")
            self.cbSystemParam.SelectedIndex = 0
            self.cbSystemParam.Enabled = False
            self.cbSystemValue.Enabled = False
            return

        for name in self.system_param_names:
            self.cbSystemParam.Items.Add(name)

        self.cbSystemParam.Items.Insert(0, u"(без фильтра по системе)")
        self.cbSystemParam.SelectedIndex = 0
        self.current_system_param = None

        self.cbSystemValue.Items.Clear()
        self.cbSystemValue.Items.Add(u"(любое значение)")
        self.cbSystemValue.SelectedIndex = 0

    def on_system_param_changed(self, sender, args):
        idx = self.cbSystemParam.SelectedIndex
        if idx <= 0:
            self.current_system_param = None
            self.cbSystemValue.Items.Clear()
            self.cbSystemValue.Items.Add(u"(любое значение)")
            self.cbSystemValue.SelectedIndex = 0
            return

        name = self.system_param_names[idx - 1]
        self.current_system_param = name

        view = self.doc.ActiveView
        collector = (
            self.DB.FilteredElementCollector(self.doc, view.Id)
            .WhereElementIsNotElementType()
        )

        values = set()
        for elem in collector:
            try:
                p = elem.LookupParameter(name)
                if p and p.StorageType == self.DB.StorageType.String:
                    val = p.AsString()
                    if val and val.strip():
                        values.add(val.strip())
            except Exception:
                continue

        sorted_vals = sorted(values)

        self.cbSystemValue.Items.Clear()
        self.cbSystemValue.Items.Add(u"(любое значение)")
        for v in sorted_vals:
            self.cbSystemValue.Items.Add(v)
        self.cbSystemValue.SelectedIndex = 0

    def get_selected_system_value(self):
        if not self.current_system_param:
            return None, None

        idx = self.cbSystemValue.SelectedIndex
        if idx <= 0:
            # Выбран параметр системы, но значение "(любое значение)" - возвращаем параметр без значения
            return self.current_system_param, None

        val = self.cbSystemValue.SelectedItem
        return self.current_system_param, val

    # ---------- Поиск по элементам в спецификации ----------

    def _collect_matching_elements(self, search_text):
        view = self.doc.ActiveView
        collector = (
            self.DB.FilteredElementCollector(self.doc, view.Id)
            .WhereElementIsNotElementType()
        )

        try:
            search_lower = search_text.lower()
        except Exception:
            search_lower = search_text

        selected_info = self.get_selected_filter_info()
        sys_param_name, sys_value = self.get_selected_system_value()

        matches = []

        for elem in collector:
            # Фильтрация по системному параметру
            if sys_param_name:
                try:
                    sp = elem.LookupParameter(sys_param_name)
                    sval = (
                        sp.AsString()
                        if sp and sp.StorageType == self.DB.StorageType.String
                        else None
                    )
                except Exception:
                    sval = None
                # Если выбрано конкретное значение - проверяем его
                if sys_value:
                    if sval != sys_value:
                        continue
                # Если значение не выбрано - пропускаем элементы без этого параметра
                else:
                    if not sval:
                        continue

            if search_lower:
                match_text = False

                # Если выбрано конкретное поле - ищем ТОЛЬКО по нему
                if selected_info is not None:
                    param_name = selected_info["title"]
                    # Проверяем параметр экземпляра
                    try:
                        p = elem.LookupParameter(param_name)
                        if p and p.StorageType == self.DB.StorageType.String:
                            val = p.AsString()
                            if val and search_lower in val.lower():
                                match_text = True
                    except Exception:
                        pass
                    
                    # Если не нашли в экземпляре - проверяем параметр типа
                    if not match_text:
                        try:
                            elem_type = self.doc.GetElement(elem.GetTypeId())
                            if elem_type:
                                p = elem_type.LookupParameter(param_name)
                                if p and p.StorageType == self.DB.StorageType.String:
                                    val = p.AsString()
                                    if val and search_lower in val.lower():
                                        match_text = True
                        except Exception:
                            pass
                else:
                    # Поиск по всем параметрам (выбрано "По всем полям")
                    for param in elem.Parameters:
                        try:
                            if param.StorageType != self.DB.StorageType.String:
                                continue
                            val = param.AsString()
                            if not val:
                                continue
                            if search_lower in val.lower():
                                match_text = True
                                break
                        except Exception:
                            continue

                if not match_text:
                    continue

            matches.append(elem)

        return matches

    def mark_found_elements(self, elements):
        if not elements:
            return

        t = self.DB.Transaction(self.doc, self.TRANSACTION_TITLE)
        try:
            t.Start()
            for elem in elements:
                try:
                    p = elem.LookupParameter(self.FOUND_PARAM_NAME)
                    if p and (not p.IsReadOnly) and p.StorageType == self.DB.StorageType.String:
                        if elem.Id not in self.original_positions:
                            self.original_positions[elem.Id] = p.AsString()
                        current_val = p.AsString()
                        if current_val != self.FOUND_PARAM_VALUE:
                            p.Set(self.FOUND_PARAM_VALUE)
                except Exception:
                    continue
            t.Commit()
        except Exception:
            try:
                t.RollBack()
            except Exception:
                pass

    def restore_previous_marks(self):
        if not self.original_positions:
            return

        t = self.DB.Transaction(self.doc, self.RESTORE_TITLE)
        try:
            t.Start()
            for eid, val in list(self.original_positions.items()):
                try:
                    elem = self.doc.GetElement(eid)
                    if not elem:
                        continue
                    p = elem.LookupParameter(self.FOUND_PARAM_NAME)
                    if p and (not p.IsReadOnly) and p.StorageType == self.DB.StorageType.String:
                        if val is None:
                            p.Set(u"")
                        else:
                            p.Set(val)
                except Exception:
                    continue
            t.Commit()
        except Exception:
            try:
                t.RollBack()
            except Exception:
                pass
        self.original_positions.clear()

    # ---------- Действия ----------

    def run_search(self):
        view = self.doc.ActiveView
        if not isinstance(view, self.DB.ViewSchedule):
            WF.MessageBox.Show(
                u"Не выбран вид спецификации (ViewSchedule).",
                u"Поиск по спецификации",
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Warning,
            )
            return

        search_text = self.tbSearch.Text or u""
        sys_param_name, sys_value = self.get_selected_system_value()
        
        # Если нет ни текста поиска, ни фильтра по системе - выходим
        if not search_text.strip() and not sys_param_name:
            return

        self.restore_previous_marks()

        matches = self._collect_matching_elements(search_text)
        if not matches:
            if search_text.strip():
                msg = u"Ничего не найдено по запросу: '{0}'.".format(search_text)
            else:
                msg = u"Элементы для выбранной системы не найдены."
            WF.MessageBox.Show(
                msg,
                u"Поиск по спецификации",
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Information,
            )
            return

        self.mark_found_elements(matches)

    def on_search_click(self, sender, args):
        if hasattr(self, "ext_handler") and self.ext_handler.has_pending_action:
            return
        self.ext_handler.set_action("search")
        self.ext_event.Raise()

    def run_select(self):
        view = self.doc.ActiveView
        if not isinstance(view, self.DB.ViewSchedule):
            WF.MessageBox.Show(
                u"Не выбран вид спецификации (ViewSchedule).",
                u"Поиск по спецификации",
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Warning,
            )
            return

        search_text = self.tbSearch.Text or u""
        sys_param_name, sys_value = self.get_selected_system_value()
        
        # Если нет ни текста поиска, ни фильтра по системе - выходим
        if not search_text.strip() and not sys_param_name:
            return

        self.restore_previous_marks()

        matches = self._collect_matching_elements(search_text)
        if not matches:
            WF.MessageBox.Show(
                u"Элементы, подходящие под условия, не найдены.",
                u"Поиск по спецификации",
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Information,
            )
            return

        self.mark_found_elements(matches)
        sel_ids = List[self.DB.ElementId]([elem.Id for elem in matches])
        self.uidoc.Selection.SetElementIds(sel_ids)
        
        # Получить или создать 3D вид пользователя и подрезать по элементам
        self._show_elements_in_3d_view(matches)

    def _get_or_create_user_3d_view(self):
        """Получить или создать 3D вид пользователя {3D - Username}."""
        username = self.doc.Application.Username
        view_name = u"{{3D - {0}}}".format(username)
        
        # Ищем существующий 3D вид пользователя
        collector = self.DB.FilteredElementCollector(self.doc).OfClass(self.DB.View3D)
        for view in collector:
            try:
                if view.Name == view_name:
                    return view
            except Exception:
                continue
        
        # Создаём новый 3D вид
        view_family_types = self.DB.FilteredElementCollector(self.doc).OfClass(self.DB.ViewFamilyType)
        view3d_type = None
        for vft in view_family_types:
            try:
                if vft.ViewFamily == self.DB.ViewFamily.ThreeDimensional:
                    view3d_type = vft
                    break
            except Exception:
                continue
        
        if not view3d_type:
            return None
        
        new_view = self.DB.View3D.CreateIsometric(self.doc, view3d_type.Id)
        try:
            new_view.Name = view_name
        except Exception:
            pass  # Имя может быть занято
        
        return new_view

    def _show_elements_in_3d_view(self, elements):
        """Открыть 3D вид и подрезать Section Box по найденным элементам."""
        if not elements:
            return
        
        t = self.DB.Transaction(self.doc, u"Подрезка 3D вида")
        try:
            t.Start()
            
            view3d = self._get_or_create_user_3d_view()
            if not view3d:
                t.RollBack()
                return
            
            # Вычисляем BoundingBox по всем элементам
            min_pt = None
            max_pt = None
            
            for elem in elements:
                try:
                    bb = elem.get_BoundingBox(None)
                    if bb is None:
                        continue
                    
                    if min_pt is None:
                        min_pt = self.DB.XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z)
                        max_pt = self.DB.XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z)
                    else:
                        min_pt = self.DB.XYZ(
                            min(min_pt.X, bb.Min.X),
                            min(min_pt.Y, bb.Min.Y),
                            min(min_pt.Z, bb.Min.Z)
                        )
                        max_pt = self.DB.XYZ(
                            max(max_pt.X, bb.Max.X),
                            max(max_pt.Y, bb.Max.Y),
                            max(max_pt.Z, bb.Max.Z)
                        )
                except Exception:
                    continue
            
            if min_pt is None or max_pt is None:
                t.RollBack()
                return
            
            # Добавляем небольшой отступ
            offset = 1.0  # ~0.3 метра
            min_pt = self.DB.XYZ(min_pt.X - offset, min_pt.Y - offset, min_pt.Z - offset)
            max_pt = self.DB.XYZ(max_pt.X + offset, max_pt.Y + offset, max_pt.Z + offset)
            
            # Создаём BoundingBoxXYZ для Section Box
            section_box = self.DB.BoundingBoxXYZ()
            section_box.Min = min_pt
            section_box.Max = max_pt
            
            # Применяем Section Box
            view3d.IsSectionBoxActive = True
            view3d.SetSectionBox(section_box)
            
            t.Commit()
            
            # Переключаемся на 3D вид
            self.uidoc.ActiveView = view3d
            
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass

    def on_select_click(self, sender, args):
        if hasattr(self, "ext_handler") and self.ext_handler.has_pending_action:
            return
        self.ext_handler.set_action("select")
        self.ext_event.Raise()

    def on_form_closing(self, sender, args):
        if self._is_closing_via_handler:
            return
        try:
            args.Cancel = True
        except Exception:
            pass
        self.ext_handler.set_action("close")
        self.ext_event.Raise()

    def perform_close_from_handler(self):
        if self._is_closing_via_handler:
            return
        self._is_closing_via_handler = True
        try:
            self.restore_previous_marks()
        except Exception:
            pass
        try:
            self.Close()
        except Exception:
            pass

    def on_close_click(self, sender, args):
        if hasattr(self, "ext_handler") and self.ext_handler.has_pending_action:
            return
        self.ext_handler.set_action("close")
        self.ext_event.Raise()

    def OnFormClosed(self, args):
        try:
            self.restore_previous_marks()
        except Exception:
            pass
        try:
            global search_form_instance
            search_form_instance = None
        except Exception:
            pass
        try:
            Form.OnFormClosed(self, args)
        except Exception:
            pass


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    view = doc.ActiveView

    if not isinstance(view, DB.ViewSchedule):
        WF.MessageBox.Show(
            u"Не выбран вид спецификации (ViewSchedule).",
            u"Поиск по спецификации",
            WF.MessageBoxButtons.OK,
            WF.MessageBoxIcon.Warning,
        )
        return

    global search_form_instance
    try:
        if search_form_instance and (not search_form_instance.IsDisposed):
            search_form_instance.Close()
    except Exception:
        pass

    search_form_instance = SearchInScheduleForm(doc, uidoc)
    search_form_instance.Show()  # modeless to keep schedule scrollable


if __name__ == "__main__":
    main()
