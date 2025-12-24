# -*- coding: utf-8 -*-
"""
Поиск выделенного элемента в спецификации.
- Если открыта одна спецификация — поиск сразу в ней
- Shift+Click — открыть окно выбора спецификации
"""
from pyrevit import revit, DB

import System.Windows.Forms as WF
from System.Windows.Forms import (
    Form, Button, Label, DialogResult, Keys, Control,
    AnchorStyles, FormStartPosition, ComboBox, ComboBoxStyle
)
from System.Drawing import Size, Point
from System.Collections.Generic import List


class FindInScheduleForm(Form):
    """Поиск выделенного семейства в выбранной ведомости
    и временная запись значения в ADSK_Позиция или ADSK_Примечание.
    """

    FOUND_PARAM_NAME = u'ADSK_Позиция'
    FOUND_PARAM_NAME_FALLBACK = u'ADSK_Примечание'
    FOUND_PARAM_VALUE = u'Найдено'

    def __init__(self, doc, uidoc):
        Form.__init__(self)

        self.doc = doc
        self.uidoc = uidoc
        self.DB = DB

        # Для передачи результата наружу
        self._schedule_to_open = None
        self._element_to_select = None

        # ---------- Свойства формы ----------
        self.Text = u'Найти в ведомости по ADSK_Позиция'
        self.StartPosition = FormStartPosition.CenterScreen
        self.Size = Size(520, 140)
        self.FormBorderStyle = WF.FormBorderStyle.FixedDialog
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.ShowInTaskbar = False
        self.TopMost = True

        # ---------- Элементы управления ----------
        self.lblSchedule = Label()
        self.lblSchedule.Text = u'Ведомость:'
        self.lblSchedule.AutoSize = True
        self.lblSchedule.Location = Point(10, 15)

        self.cbSchedule = ComboBox()
        self.cbSchedule.DropDownStyle = ComboBoxStyle.DropDownList
        self.cbSchedule.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.cbSchedule.Location = Point(90, 10)
        self.cbSchedule.Width = 400

        self.btnMark = Button()
        self.btnMark.Text = u'Найти и отметить'
        self.btnMark.AutoSize = True
        self.btnMark.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.btnClear = Button()
        self.btnClear.Text = u'Очистить "Найдено"'
        self.btnClear.AutoSize = True
        self.btnClear.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.btnClose = Button()
        self.btnClose.Text = u'Закрыть'
        self.btnClose.AutoSize = True
        self.btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        # Добавление контролов
        self.Controls.Add(self.lblSchedule)
        self.Controls.Add(self.cbSchedule)
        self.Controls.Add(self.btnMark)
        self.Controls.Add(self.btnClear)
        self.Controls.Add(self.btnClose)

        # Список ведомостей
        self._schedule_ids = []
        self.build_schedule_list()

        # События
        self.btnMark.Click += self.on_mark_click
        self.btnClear.Click += self.on_clear_click
        self.btnClose.Click += self.on_close_click
        self.cbSchedule.SelectedIndexChanged += self.on_schedule_changed

        self._update_bottom_buttons_positions()

    # ---------- Служебные методы ----------
    def _update_bottom_buttons_positions(self):
        margin_bottom = 35
        margin_right = 20
        bottom_y = self.ClientSize.Height - margin_bottom

        self.btnClose.Location = Point(
            self.ClientSize.Width - self.btnClose.Width - margin_right,
            bottom_y
        )
        self.btnClear.Location = Point(
            self.btnClose.Left - self.btnClear.Width - 10,
            bottom_y
        )
        self.btnMark.Location = Point(
            self.btnClear.Left - self.btnMark.Width - 10,
            bottom_y
        )

    def OnResize(self, args):
        Form.OnResize(self, args)
        try:
            self._update_bottom_buttons_positions()
        except Exception:
            pass

    def on_schedule_changed(self, sender, args):
        idx = self.cbSchedule.SelectedIndex
        if idx < 0 or idx >= len(self._schedule_ids):
            self._schedule_id = None
        else:
            self._schedule_id = self._schedule_ids[idx]

    def build_schedule_list(self):
        self.cbSchedule.Items.Clear()
        self._schedule_ids = []

        collector = DB.FilteredElementCollector(self.doc).OfClass(DB.ViewSchedule)

        default_index = -1
        for vs in collector:
            try:
                if vs.IsTemplate:
                    continue
            except Exception:
                continue

            name = vs.Name
            idx = self.cbSchedule.Items.Add(name)
            self._schedule_ids.append(vs.Id)

            if (default_index < 0 and
                    name and u'Форма 1_ГОСТ 21.110-2013' in name):
                default_index = idx

        if self.cbSchedule.Items.Count > 0:
            if default_index >= 0:
                self.cbSchedule.SelectedIndex = default_index
            else:
                self.cbSchedule.SelectedIndex = 0
            self._schedule_id = self._schedule_ids[self.cbSchedule.SelectedIndex]
        else:
            WF.MessageBox.Show(
                u'В проекте не найдено ни одной ведомости.',
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Warning
            )

    def get_selected_schedule(self):
        idx = self.cbSchedule.SelectedIndex
        if idx < 0 or idx >= len(self._schedule_ids):
            return None
        sid = self._schedule_ids[idx]
        view = self.doc.GetElement(sid)
        if isinstance(view, DB.ViewSchedule):
            return view
        return None

    # ---------- Работа с параметром ----------
    def _ensure_found_param(self, element):
        """Ищет подходящий параметр для записи. Сначала ADSK_Позиция, затем ADSK_Примечание.
        Проверяет как параметры экземпляра, так и параметры типа.
        Возвращает (параметр, элемент/тип, is_type_param)."""
        
        # Сначала пробуем основной параметр экземпляра
        p = element.LookupParameter(self.FOUND_PARAM_NAME)
        if (p is not None and
                p.StorageType == DB.StorageType.String and
                not p.IsReadOnly):
            return p, element, False
        
        # Пробуем основной параметр типа
        try:
            elem_type = self.doc.GetElement(element.GetTypeId())
            if elem_type:
                p = elem_type.LookupParameter(self.FOUND_PARAM_NAME)
                if (p is not None and
                        p.StorageType == DB.StorageType.String and
                        not p.IsReadOnly):
                    return p, elem_type, True
        except Exception:
            pass
        
        # Если не найден - пробуем fallback параметр экземпляра
        p = element.LookupParameter(self.FOUND_PARAM_NAME_FALLBACK)
        if (p is not None and
                p.StorageType == DB.StorageType.String and
                not p.IsReadOnly):
            return p, element, False
        
        # Пробуем fallback параметр типа
        try:
            elem_type = self.doc.GetElement(element.GetTypeId())
            if elem_type:
                p = elem_type.LookupParameter(self.FOUND_PARAM_NAME_FALLBACK)
                if (p is not None and
                        p.StorageType == DB.StorageType.String and
                        not p.IsReadOnly):
                    return p, elem_type, True
        except Exception:
            pass
        
        return None, None, False

    def get_grouping_field_ids(self, schedule):
        """Получает список ScheduleFieldId для полей, по которым идёт группировка/сортировка."""
        grouping_field_ids = []
        try:
            definition = schedule.Definition
            if definition is None:
                return grouping_field_ids
            
            sort_group_fields = definition.GetSortGroupFields()
            if sort_group_fields:
                for sgf in sort_group_fields:
                    grouping_field_ids.append(sgf.FieldId)
        except Exception:
            pass
        return grouping_field_ids

    def get_param_value_for_field(self, element, field, schedule):
        """Получает значение параметра элемента, соответствующего полю спецификации."""
        try:
            param_id = field.ParameterId
            if param_id == DB.ElementId.InvalidElementId:
                return None
            
            field_name = field.GetName()
            param = element.LookupParameter(field_name)
            
            if param is None:
                try:
                    bip = param_id.IntegerValue
                    param = element.get_Parameter(DB.BuiltInParameter(bip))
                except Exception:
                    pass
            
            if param is None:
                return None
            
            storage_type = param.StorageType
            if storage_type == DB.StorageType.String:
                return param.AsString() or ''
            elif storage_type == DB.StorageType.Integer:
                return param.AsInteger()
            elif storage_type == DB.StorageType.Double:
                return param.AsDouble()
            elif storage_type == DB.StorageType.ElementId:
                return param.AsElementId().IntegerValue
            else:
                return param.AsValueString() or ''
        except Exception:
            return None

    def get_element_grouping_key(self, element, grouping_fields, schedule):
        """Создаёт ключ группировки для элемента на основе значений полей группировки."""
        key_parts = []
        for field in grouping_fields:
            val = self.get_param_value_for_field(element, field, schedule)
            key_parts.append(str(val) if val is not None else '')
        return tuple(key_parts)

    def find_elements_in_same_row(self, target_element, schedule):
        """Находит все элементы в спецификации, которые находятся в той же строке."""
        result = [target_element]
        
        try:
            definition = schedule.Definition
            if definition is None:
                return result
            
            grouping_field_ids = self.get_grouping_field_ids(schedule)
            
            if not grouping_field_ids:
                return result
            
            grouping_fields = []
            for fid in grouping_field_ids:
                try:
                    field = definition.GetField(fid)
                    if field is not None:
                        grouping_fields.append(field)
                except Exception:
                    pass
            
            if not grouping_fields:
                return result
            
            target_key = self.get_element_grouping_key(target_element, grouping_fields, schedule)
            
            collector = DB.FilteredElementCollector(self.doc, schedule.Id) \
                          .WhereElementIsNotElementType()
            
            for el in collector:
                if el.Id == target_element.Id:
                    continue
                try:
                    el_key = self.get_element_grouping_key(el, grouping_fields, schedule)
                    if el_key == target_key:
                        result.append(el)
                except Exception:
                    continue
        except Exception:
            pass
        
        return result

    def mark_elements(self, elements):
        """Записывает значение в параметр ADSK_Позиция для списка элементов. Возвращает True при успехе."""
        if not elements:
            return False
        
        elements_with_param = []
        used_type_param = False
        for el in elements:
            p, target_elem, is_type = self._ensure_found_param(el)
            if p is not None:
                elements_with_param.append((target_elem, p))
                if is_type:
                    used_type_param = True
        
        if not elements_with_param:
            WF.MessageBox.Show(
                u'У элементов нет доступного текстового параметра "{0}" или "{1}".'
                .format(self.FOUND_PARAM_NAME, self.FOUND_PARAM_NAME_FALLBACK),
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Warning
            )
            return False

        t = DB.Transaction(self.doc, u'Установить {0}'.format(self.FOUND_PARAM_NAME))
        try:
            t.Start()
            success_count = 0
            for el, p in elements_with_param:
                try:
                    result = p.Set(self.FOUND_PARAM_VALUE)
                    if result:
                        success_count += 1
                except Exception:
                    pass
            
            if success_count == 0:
                t.RollBack()
                WF.MessageBox.Show(
                    u'Не удалось записать значение в параметр "{0}".\n'
                    u'Возможно, параметр доступен только для чтения.'.format(self.FOUND_PARAM_NAME),
                    u'Найти в ведомости',
                    WF.MessageBoxButtons.OK,
                    WF.MessageBoxIcon.Warning
                )
                return False
            t.Commit()
            
            # Предупреждение если использован параметр типа
            if used_type_param:
                WF.MessageBox.Show(
                    u'Внимание: использован параметр типа элемента.\n\n'
                    u'Если в спецификации включена группировка, отображение\n'
                    u'метки "Найдено" может быть некорректным — она будет\n'
                    u'показана для всех элементов данного типа.',
                    u'Найти в ведомости',
                    WF.MessageBoxButtons.OK,
                    WF.MessageBoxIcon.Information
                )
            
            return True
        except Exception as exc:
            try:
                t.RollBack()
            except Exception:
                pass
            WF.MessageBox.Show(
                u'Ошибка при записи параметра:\n{0}'.format(exc),
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Error
            )
            return False

    # ---------- Основная логика ----------
    def perform_mark(self, schedule=None):
        """Выполняет поиск и отметку элемента. schedule можно передать напрямую."""
        # 1. Выбранный элемент
        sel_ids = list(self.uidoc.Selection.GetElementIds())
        if len(sel_ids) != 1:
            WF.MessageBox.Show(
                u'Выберите один элемент семейства в модели перед запуском команды.',
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Information
            )
            return False

        element = self.doc.GetElement(sel_ids[0])
        if element is None:
            WF.MessageBox.Show(
                u'Не удалось получить выбранный элемент.',
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Error
            )
            return False

        # 2. Выбранная ведомость
        if schedule is None:
            schedule = self.get_selected_schedule()
        if schedule is None:
            WF.MessageBox.Show(
                u'Выберите ведомость.',
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Warning
            )
            return False

        # 3. Проверяем, входит ли элемент в эту ведомость
        try:
            collector = DB.FilteredElementCollector(self.doc, schedule.Id) \
                          .WhereElementIsNotElementType()
        except Exception as exc:
            WF.MessageBox.Show(
                u'Ошибка при получении элементов из ведомости:\n{0}'.format(exc),
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Error
            )
            return False

        found = False
        if collector is not None:
            for el in collector:
                try:
                    if el.Id == element.Id:
                        found = True
                        break
                except Exception:
                    continue

        if not found:
            WF.MessageBox.Show(
                u'Выбранный элемент не найден в выбранной ведомости.',
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Information
            )
            return False

        # 4. Находим все элементы в той же строке спецификации
        elements_in_row = self.find_elements_in_same_row(element, schedule)

        # 5. Записываем ADSK_Позиция = "Найдено" у всех элементов строки
        mark_success = self.mark_elements(elements_in_row)
        if not mark_success:
            return False

        # 6. Сохраняем schedule для открытия после закрытия формы
        self._schedule_to_open = schedule
        self._element_to_select = element

        return True

    # ---------- Обработчики событий ----------
    def on_mark_click(self, sender, args):
        try:
            success = self.perform_mark()
            if success:
                self.DialogResult = DialogResult.OK
                self.Close()
        except Exception as exc:
            WF.MessageBox.Show(
                u'Ошибка:\n{0}'.format(exc),
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Error
            )

    def on_clear_click(self, sender, args):
        """Очищает значение 'Найдено' у всех элементов в проекте."""
        try:
            clear_found_marks(self.doc)
            WF.MessageBox.Show(
                u'Значения "{0}" очищены.'.format(self.FOUND_PARAM_VALUE),
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Information
            )
        except Exception as exc:
            WF.MessageBox.Show(
                u'Ошибка при очистке:\n{0}'.format(exc),
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Error
            )

    def on_close_click(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


def get_open_schedule_views(uidoc):
    """Возвращает список открытых спецификаций (ViewSchedule)."""
    open_schedules = []
    try:
        open_views = uidoc.GetOpenUIViews()
        for uiview in open_views:
            view = uidoc.Document.GetElement(uiview.ViewId)
            if isinstance(view, DB.ViewSchedule) and not view.IsTemplate:
                open_schedules.append(view)
    except Exception:
        pass
    return open_schedules


def clear_found_marks(doc, found_value=u'Найдено', param_names=None):
    """Очищает параметры ADSK_Позиция и ADSK_Примечание у всех элементов и типов, где значение = 'Найдено'."""
    if param_names is None:
        param_names = [u'ADSK_Позиция', u'ADSK_Примечание']
    
    elements_to_clear = []
    
    # Проверяем экземпляры элементов
    collector = DB.FilteredElementCollector(doc) \
                  .WhereElementIsNotElementType() \
                  .ToElements()
    
    for el in collector:
        for param_name in param_names:
            try:
                p = el.LookupParameter(param_name)
                if p is None or p.IsReadOnly:
                    continue
                if p.StorageType != DB.StorageType.String:
                    continue
                val = p.AsString()
                if val == found_value:
                    elements_to_clear.append((el, p))
            except Exception:
                continue
    
    # Проверяем типы элементов
    type_collector = DB.FilteredElementCollector(doc) \
                       .WhereElementIsElementType() \
                       .ToElements()
    
    for el_type in type_collector:
        for param_name in param_names:
            try:
                p = el_type.LookupParameter(param_name)
                if p is None or p.IsReadOnly:
                    continue
                if p.StorageType != DB.StorageType.String:
                    continue
                val = p.AsString()
                if val == found_value:
                    elements_to_clear.append((el_type, p))
            except Exception:
                continue
    
    if not elements_to_clear:
        return
    
    t = DB.Transaction(doc, u'Очистить метки Найдено')
    try:
        t.Start()
        for el, p in elements_to_clear:
            try:
                p.Set('')
            except Exception:
                pass
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass


def perform_quick_search(doc, uidoc, schedule):
    """Выполняет быстрый поиск без открытия формы."""
    FOUND_PARAM_NAME = u'ADSK_Позиция'
    FOUND_PARAM_NAME_FALLBACK = u'ADSK_Примечание'
    FOUND_PARAM_VALUE = u'Найдено'
    
    # 1. Очищаем предыдущие метки "Найдено"
    clear_found_marks(doc)
    
    # 2. Получаем выбранный элемент
    sel_ids = list(uidoc.Selection.GetElementIds())
    if len(sel_ids) != 1:
        WF.MessageBox.Show(
            u'Выберите один элемент семейства в модели.',
            u'Найти в ведомости',
            WF.MessageBoxButtons.OK,
            WF.MessageBoxIcon.Information
        )
        return False

    element = doc.GetElement(sel_ids[0])
    if element is None:
        WF.MessageBox.Show(
            u'Не удалось получить выбранный элемент.',
            u'Найти в ведомости',
            WF.MessageBoxButtons.OK,
            WF.MessageBoxIcon.Error
        )
        return False

    # 3. Проверяем, входит ли элемент в эту ведомость
    try:
        collector = DB.FilteredElementCollector(doc, schedule.Id) \
                      .WhereElementIsNotElementType()
    except Exception as exc:
        WF.MessageBox.Show(
            u'Ошибка при получении элементов из ведомости:\n{0}'.format(exc),
            u'Найти в ведомости',
            WF.MessageBoxButtons.OK,
            WF.MessageBoxIcon.Error
        )
        return False

    found = False
    for el in collector:
        try:
            if el.Id == element.Id:
                found = True
                break
        except Exception:
            continue

    if not found:
        WF.MessageBox.Show(
            u'Выбранный элемент не найден в открытой ведомости "{0}".'.format(schedule.Name),
            u'Найти в ведомости',
            WF.MessageBoxButtons.OK,
            WF.MessageBoxIcon.Information
        )
        return False

    # 4. Находим все элементы в той же строке спецификации
    # Создаём временную форму для использования её методов
    temp_form = FindInScheduleForm(doc, uidoc)
    elements_in_row = temp_form.find_elements_in_same_row(element, schedule)
    
    # 5. Записываем ADSK_Позиция или ADSK_Примечание = "Найдено"
    # Проверяем экземпляр и тип элемента
    elements_with_param = []
    used_type_param = False
    for el in elements_in_row:
        # Проверяем параметр экземпляра
        p = el.LookupParameter(FOUND_PARAM_NAME)
        if p is not None and not p.IsReadOnly and p.StorageType == DB.StorageType.String:
            elements_with_param.append((el, p))
            continue
        
        # Проверяем параметр типа
        try:
            elem_type = doc.GetElement(el.GetTypeId())
            if elem_type:
                p = elem_type.LookupParameter(FOUND_PARAM_NAME)
                if p is not None and not p.IsReadOnly and p.StorageType == DB.StorageType.String:
                    elements_with_param.append((elem_type, p))
                    used_type_param = True
                    continue
        except Exception:
            pass
        
        # Пробуем fallback параметр экземпляра
        p = el.LookupParameter(FOUND_PARAM_NAME_FALLBACK)
        if p is not None and not p.IsReadOnly and p.StorageType == DB.StorageType.String:
            elements_with_param.append((el, p))
            continue
        
        # Пробуем fallback параметр типа
        try:
            elem_type = doc.GetElement(el.GetTypeId())
            if elem_type:
                p = elem_type.LookupParameter(FOUND_PARAM_NAME_FALLBACK)
                if p is not None and not p.IsReadOnly and p.StorageType == DB.StorageType.String:
                    elements_with_param.append((elem_type, p))
                    used_type_param = True
        except Exception:
            pass
    
    if not elements_with_param:
        WF.MessageBox.Show(
            u'У элементов нет доступного текстового параметра "{0}" или "{1}".'.format(
                FOUND_PARAM_NAME, FOUND_PARAM_NAME_FALLBACK),
            u'Найти в ведомости',
            WF.MessageBoxButtons.OK,
            WF.MessageBoxIcon.Warning
        )
        return False

    t = DB.Transaction(doc, u'Установить {0}'.format(FOUND_PARAM_NAME))
    try:
        t.Start()
        for el, p in elements_with_param:
            try:
                p.Set(FOUND_PARAM_VALUE)
            except Exception:
                pass
        t.Commit()
        
        # Предупреждение если использован параметр типа
        if used_type_param:
            WF.MessageBox.Show(
                u'Внимание: использован параметр типа элемента.\n\n'
                u'Если в спецификации включена группировка, отображение\n'
                u'метки "Найдено" может быть некорректным — она будет\n'
                u'показана для всех элементов данного типа.',
                u'Найти в ведомости',
                WF.MessageBoxButtons.OK,
                WF.MessageBoxIcon.Information
            )
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return False

    # 6. Переключаемся на ведомость
    try:
        uidoc.ActiveView = schedule
    except Exception:
        try:
            uidoc.RequestViewChange(schedule)
        except Exception:
            pass

    # 7. Выделяем элемент
    try:
        uidoc.Selection.SetElementIds(List[DB.ElementId]([element.Id]))
    except Exception:
        pass

    return True


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    
    # Проверяем, нажат ли Shift
    shift_pressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift
    
    # Получаем открытые спецификации
    open_schedules = get_open_schedule_views(uidoc)
    
    # Если открыта ровно одна спецификация и Shift не нажат — быстрый поиск
    if len(open_schedules) == 1 and not shift_pressed:
        # Очищаем и ищем сразу
        perform_quick_search(doc, uidoc, open_schedules[0])
    else:
        # Открываем форму выбора
        # Сначала очищаем предыдущие метки
        clear_found_marks(doc)
        
        form = FindInScheduleForm(doc, uidoc)
        result = form.ShowDialog()

        # Если нажали "Найти и отметить" успешно
        if result == DialogResult.OK:
            schedule = getattr(form, '_schedule_to_open', None)
            element = getattr(form, '_element_to_select', None)
            
            if schedule is not None:
                try:
                    uidoc.ActiveView = schedule
                except Exception:
                    try:
                        uidoc.RequestViewChange(schedule)
                    except Exception:
                        pass

            if element is not None:
                try:
                    uidoc.Selection.SetElementIds(List[DB.ElementId]([element.Id]))
                except Exception:
                    pass


if __name__ == '__main__':
    main()
