# -*- coding: utf-8 -*-
# SuperFilter (3D/планы/разрезы) для pyRevit (IronPython)
# Версия: WinForms v2.7
from __future__ import print_function, division

import clr
from collections import OrderedDict, defaultdict

from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementId, StorageType, View,
    Category, CategoryType, BuiltInCategory, BuiltInParameter, ParameterElement
)
from Autodesk.Revit.Exceptions import InvalidOperationException
from pyrevit import revit, script

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

# ---------------- Получаем выделенный элемент ----------------
def get_preselected_element():
    """Получить первый выделенный элемент перед запуском плагина"""
    try:
        sel_ids = uidoc.Selection.GetElementIds()
        if sel_ids and sel_ids.Count > 0:
            for eid in sel_ids:
                el = doc.GetElement(eid)
                if el and el.Category and el.Category.CategoryType == CategoryType.Model:
                    return el
    except:
        pass
    return None

def get_element_category_name(el):
    """Получить имя категории элемента"""
    try:
        if el and el.Category:
            return _to_unicode(el.Category.Name)
    except:
        pass
    return None

def get_element_family_and_type(el):
    """Получить строку 'Семейство и типоразмер' элемента"""
    try:
        # Ищем параметр "Семейство и типоразмер" (ELEM_FAMILY_AND_TYPE_PARAM)
        p = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
        if p and p.HasValue:
            val = p.AsValueString()
            if val:
                return _to_unicode(val)
        # Альтернативный способ - через тип
        tid = el.GetTypeId()
        if tid and tid.IntegerValue != -1:
            typ = doc.GetElement(tid)
            if typ:
                # Пытаемся получить имя семейства и типа
                fam_name = None
                type_name = None
                try:
                    fam_name = typ.FamilyName
                except:
                    pass
                try:
                    type_name = typ.Name
                except:
                    pass
                if fam_name and type_name:
                    return u'{0}: {1}'.format(_to_unicode(fam_name), _to_unicode(type_name))
                elif type_name:
                    return _to_unicode(type_name)
    except:
        pass
    return None

# Сохраняем выделенный элемент до начала работы
_preselected_element = get_preselected_element()

# ---------------- Helpers ----------------
OST_CAMERAS_INT = int(BuiltInCategory.OST_Cameras)
_PROJECT_DEF_NAMES = None  # кэш имён определений параметров проекта

def ensure_view_supported(view):
    if view is None or view.IsTemplate:
        from pyrevit import forms
        forms.alert(u'Активируйте обычный вид (не шаблон) и запустите снова.', title=u'Суперфильтр')
        raise SystemExit

def _to_unicode(s):
    try:
        return unicode(s)
    except:
        try:
            return s.decode('utf-8')
        except:
            return s

def _is_camera(el):
    try:
        cat = el.Category
        if cat and cat.Id and cat.Id.IntegerValue == OST_CAMERAS_INT:
            return True
    except:
        pass
    return False

def _eid_to_disp(eid):
    if not isinstance(eid, ElementId):
        return _to_unicode(eid)
    # Проверяем категорию
    try:
        cat = Category.GetCategory(doc, eid)
        if cat is not None:
            return _to_unicode(cat.Name)
    except:
        pass
    # Получаем элемент
    try:
        refel = doc.GetElement(eid)
        if refel:
            # Для типоразмеров (FamilySymbol) формируем "Семейство: Тип"
            fam_name = None
            type_name = None
            try:
                fam_name = getattr(refel, 'FamilyName', None)
            except:
                pass
            try:
                type_name = refel.Name
            except:
                pass
            if fam_name and type_name:
                return u'{0}: {1}'.format(_to_unicode(fam_name), _to_unicode(type_name))
            elif type_name:
                return _to_unicode(type_name)
            elif fam_name:
                return _to_unicode(fam_name)
    except:
        pass
    return u'Id:{0}'.format(eid.IntegerValue)

def collect_candidates(only_visible=True):
    if only_visible:
        col = FilteredElementCollector(doc, active_view.Id)
    else:
        col = FilteredElementCollector(doc)
    col = col.WhereElementIsNotElementType()
    result = []
    for el in col:
        if _is_camera(el):
            continue
        cat = el.Category
        if cat and cat.CategoryType == CategoryType.Model:
            result.append(el)
    return result

def _is_builtin_param(p):
    try:
        idef = p.Definition
        bip = getattr(idef, 'BuiltInParameter', None)
        if bip is None:
            return False
        try:
            return bip != BuiltInParameter.INVALID
        except:
            try:
                return int(bip) != -1
            except:
                return False
    except:
        return False

def _is_shared_param(p):
    try:
        return bool(getattr(p, 'IsShared', False))
    except:
        return False

def _ensure_project_def_names():
    global _PROJECT_DEF_NAMES
    if _PROJECT_DEF_NAMES is not None:
        return
    names = set()
    try:
        for pe in FilteredElementCollector(doc).OfClass(ParameterElement):
            try:
                d = pe.GetDefinition()
                nm = d.Name
                if nm:
                    names.add(nm)
            except:
                pass
    except:
        pass
    _PROJECT_DEF_NAMES = names

def _is_project_param_def(defn):
    # 1) Пытаемся спросить у BindingMap напрямую
    try:
        if doc.ParameterBindings.Contains(defn):
            return True
    except:
        pass
    # 2) Fallback по кэшу имён ParameterElement'ов
    _ensure_project_def_names()
    try:
        return defn.Name in _PROJECT_DEF_NAMES
    except:
        return False

def _classify_param(p):
    # Возвращает 'builtin' | 'shared' | 'project' | 'family'
    if _is_builtin_param(p):
        return 'builtin'
    if _is_shared_param(p):
        return 'shared'
    try:
        defn = p.Definition
    except:
        defn = None
    if defn is not None and _is_project_param_def(defn):
        return 'project'
    return 'family'

def _include_param(p, include_family):
    kind = _classify_param(p)
    if kind == 'family' and not include_family:
        return False
    return True

def build_param_index(only_visible=True, include_type=True, include_family=False, sample_per_cat=30):
    elems = collect_candidates(only_visible)
    by_cat = defaultdict(list)
    for el in elems:
        try:
            cid = el.Category.Id.IntegerValue
        except:
            continue
        if len(by_cat[cid]) < sample_per_cat:
            by_cat[cid].append(el)

    param_index = {}
    type_cache = {}
    out = script.get_output()
    try:
        pb = out.create_progress_bar(len(by_cat), title=u'Подготовка параметров...')
    except:
        pb = None

    try:
        for cid, samples in by_cat.items():
            if pb: pb.update()
            for el in samples:
                for p in el.Parameters:
                    if not _include_param(p, include_family):
                        continue
                    defn = p.Definition
                    if not defn:
                        continue
                    pname = defn.Name
                    if pname not in param_index:
                        param_index[pname] = p.StorageType
                if include_type:
                    tid = el.GetTypeId()
                    if tid and tid.IntegerValue not in type_cache:
                        typ = doc.GetElement(tid)
                        type_cache[tid.IntegerValue] = True
                        if typ:
                            for p in typ.Parameters:
                                if not _include_param(p, include_family):
                                    continue
                                defn = p.Definition
                                if not defn:
                                    continue
                                pname = defn.Name
                                if pname not in param_index:
                                    param_index[pname] = p.StorageType
    finally:
        if pb: pb.close()

    names = sorted(param_index.keys(), key=lambda s: s.lower())
    return names, param_index

def collect_values_for_param(pname, storage, only_visible=True, include_type=True, include_family=False):
    elems = collect_candidates(only_visible)
    values = OrderedDict()
    has_empty_box = [False]

    def add_p(p):
        if not _include_param(p, include_family):
            return
        st = p.StorageType
        if st != storage:
            return
        try:
            if st == StorageType.String:
                sval = p.AsString()
                if (sval is None) or (sval == u''):
                    has_empty_box[0] = True
                else:
                    if sval not in values:
                        values[sval] = sval
            elif st == StorageType.Integer:
                if not getattr(p, 'HasValue', True):
                    has_empty_box[0] = True
                else:
                    ival = p.AsInteger()
                    disp = p.AsValueString() if p.AsValueString() else unicode(ival)
                    if disp not in values:
                        values[disp] = ival
            elif st == StorageType.Double:
                if not getattr(p, 'HasValue', True):
                    has_empty_box[0] = True
                else:
                    dval = p.AsDouble()
                    disp = p.AsValueString() if p.AsValueString() else unicode(dval)
                    if disp not in values:
                        values[disp] = dval
            elif st == StorageType.ElementId:
                eid = p.AsElementId()
                if eid == ElementId.InvalidElementId:
                    has_empty_box[0] = True
                else:
                    disp = _eid_to_disp(eid)
                    if disp not in values:
                        values[disp] = eid
            else:
                has_empty_box[0] = True
        except:
            pass

    out = script.get_output()
    try:
        pb = out.create_progress_bar(len(elems), title=u'Собираю значения: {0}'.format(pname))
    except:
        pb = None

    try:
        for el in elems:
            if pb: pb.update()
            p = el.LookupParameter(pname)
            if p: add_p(p)
            if include_type and (p is None or getattr(p, 'StorageType', None) != storage):
                try:
                    typ = doc.GetElement(el.GetTypeId())
                    if typ:
                        pt = typ.LookupParameter(pname)
                        if pt: add_p(pt)
                except:
                    pass
    finally:
        if pb: pb.close()

    disp_list = list(values.keys())
    try:
        disp_list.sort(key=lambda x: _to_unicode(x).lower())
    except:
        disp_list.sort()

    return disp_list, values, has_empty_box[0]

def get_param(elem, pname, lookup_in_type=True):
    p = elem.LookupParameter(pname)
    if p is None and lookup_in_type:
        try:
            et = doc.GetElement(elem.GetTypeId())
            if et:
                p = et.LookupParameter(pname)
        except:
            p = None
    return p

def match_condition(elem, pname, op, raw_value, lookup_in_type):
    p = get_param(elem, pname, lookup_in_type)
    if op == u'пусто':
        return (p is None) or (not getattr(p, 'HasValue', True)) or                (p.StorageType == StorageType.String and (p.AsString() in (None, u''))) or                (p.StorageType == StorageType.ElementId and p.AsElementId() == ElementId.InvalidElementId)
    if op == u'не пусто':
        return not match_condition(elem, pname, u'пусто', None, lookup_in_type)

    if p is None:
        return False

    st = p.StorageType
    try:
        if st == StorageType.String:
            left = p.AsString()
            if left is None:
                left = p.AsValueString()
            eq = _to_unicode(left or u'') == _to_unicode(raw_value or u'')
            return eq if op == u'=' else (not eq)
        if st == StorageType.Integer:
            eq = int(p.AsInteger()) == int(raw_value)
            return eq if op == u'=' else (not eq)
        if st == StorageType.Double:
            left = float(p.AsDouble())
            right = float(raw_value)
            eq = abs(left - right) < 1e-9
            return eq if op == u'=' else (not eq)
        if st == StorageType.ElementId:
            left = p.AsElementId()
            if not isinstance(raw_value, ElementId):
                return False
            eq = (left == raw_value)
            return eq if op == u'=' else (not eq)
    except:
        return False
    return False

# ---------------- WinForms UI ----------------
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Form, Label, ComboBox, CheckBox, Button, DialogResult,
                                 FormBorderStyle, ComboBoxStyle, AutoCompleteMode, AutoCompleteSource,
                                 Panel, BorderStyle, FlatStyle, GroupBox, AnchorStyles, FormStartPosition)
from System.Drawing import Size, Point, Color, Font, FontStyle, ContentAlignment

OPS = [u'=', u'!=', u'пусто', u'не пусто']

# Цвета в стиле референса
COLOR_ACCENT_BLUE = Color.FromArgb(0, 122, 204)      # Голубой акцент
COLOR_ACCENT_PINK = Color.FromArgb(233, 30, 99)      # Розовый акцент
COLOR_BTN_PRIMARY = Color.FromArgb(0, 150, 136)      # Бирюзовый для основных кнопок
COLOR_BTN_SECONDARY = Color.FromArgb(96, 125, 139)   # Серо-голубой для вторичных
COLOR_HEADER_BG = Color.FromArgb(245, 245, 245)      # Светлый фон заголовков
COLOR_ROW_ALT = Color.FromArgb(250, 250, 255)        # Альтернативный цвет строки

class FilterForm(Form):
    def __init__(self):
        self.Text = u'Суперфильтр'
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(900, 520)
        self.BackColor = Color.White
        self.StartPosition = FormStartPosition.CenterScreen

        # Верхняя панель настроек
        self.pnlHeader = Panel()
        self.pnlHeader.Location = Point(0, 0)
        self.pnlHeader.Size = Size(900, 80)
        self.pnlHeader.BackColor = COLOR_HEADER_BG
        self.Controls.Add(self.pnlHeader)

        y = 15
        self.lblScope = Label()
        self.lblScope.Text = u'Область поиска:'
        self.lblScope.Location = Point(15, y)
        self.lblScope.Size = Size(100, 18)
        self.lblScope.Font = Font(self.Font.FontFamily, 9, FontStyle.Bold)
        self.pnlHeader.Controls.Add(self.lblScope)

        self.cbOnlyVis = CheckBox()
        self.cbOnlyVis.Text = u'Только видимые на активном виде'
        self.cbOnlyVis.Checked = True
        self.cbOnlyVis.Location = Point(120, y-2)
        self.cbOnlyVis.Size = Size(230, 22)
        self.cbOnlyVis.BackColor = Color.Transparent
        self.cbOnlyVis.CheckedChanged += self._on_scope_changed
        self.pnlHeader.Controls.Add(self.cbOnlyVis)

        self.cbIncludeTypes = CheckBox()
        self.cbIncludeTypes.Text = u'Параметры типов'
        self.cbIncludeTypes.Checked = True
        self.cbIncludeTypes.Location = Point(360, y-2)
        self.cbIncludeTypes.Size = Size(140, 22)
        self.cbIncludeTypes.BackColor = Color.Transparent
        self.cbIncludeTypes.CheckedChanged += self._on_scope_changed
        self.pnlHeader.Controls.Add(self.cbIncludeTypes)

        self.lblLogic = Label()
        self.lblLogic.Text = u'Логика:'
        self.lblLogic.Location = Point(520, y)
        self.lblLogic.Size = Size(55, 18)
        self.lblLogic.Font = Font(self.Font.FontFamily, 9, FontStyle.Bold)
        self.pnlHeader.Controls.Add(self.lblLogic)

        self.cmbLogic = ComboBox()
        self.cmbLogic.Items.Add(u'ИЛИ')
        self.cmbLogic.Items.Add(u'И')
        self.cmbLogic.SelectedIndex = 0
        self.cmbLogic.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmbLogic.Location = Point(575, y-3)
        self.cmbLogic.Size = Size(70, 24)
        self.pnlHeader.Controls.Add(self.cmbLogic)

        y += 35
        self.btnLoadFamily = Button()
        self.btnLoadFamily.Text = u'⟳ Загрузить параметры семейства'
        self.btnLoadFamily.Location = Point(15, y)
        self.btnLoadFamily.Size = Size(250, 28)
        self.btnLoadFamily.FlatStyle = FlatStyle.Flat
        self.btnLoadFamily.BackColor = COLOR_BTN_SECONDARY
        self.btnLoadFamily.ForeColor = Color.White
        self.btnLoadFamily.FlatAppearance.BorderSize = 0
        self.btnLoadFamily.Click += self._on_load_family
        self.pnlHeader.Controls.Add(self.btnLoadFamily)

        # Группа условий фильтрации
        y = 90
        self.grpConditions = GroupBox()
        self.grpConditions.Text = u' Условия фильтрации '
        self.grpConditions.Location = Point(10, y)
        self.grpConditions.Size = Size(760, 310)
        self.grpConditions.Font = Font(self.Font.FontFamily, 9, FontStyle.Bold)
        self.Controls.Add(self.grpConditions)

        # Заголовки колонок
        innerY = 22
        lblHdrParam = Label()
        lblHdrParam.Text = u'Параметр'
        lblHdrParam.Location = Point(85, innerY)
        lblHdrParam.Size = Size(330, 18)
        lblHdrParam.Font = Font(self.Font.FontFamily, 8, FontStyle.Bold)
        lblHdrParam.ForeColor = Color.Gray
        self.grpConditions.Controls.Add(lblHdrParam)

        lblHdrOp = Label()
        lblHdrOp.Text = u'Оператор'
        lblHdrOp.Location = Point(420, innerY)
        lblHdrOp.Size = Size(100, 18)
        lblHdrOp.Font = Font(self.Font.FontFamily, 8, FontStyle.Bold)
        lblHdrOp.ForeColor = Color.Gray
        self.grpConditions.Controls.Add(lblHdrOp)

        lblHdrVal = Label()
        lblHdrVal.Text = u'Значение'
        lblHdrVal.Location = Point(535, innerY)
        lblHdrVal.Size = Size(200, 18)
        lblHdrVal.Font = Font(self.Font.FontFamily, 8, FontStyle.Bold)
        lblHdrVal.ForeColor = Color.Gray
        self.grpConditions.Controls.Add(lblHdrVal)

        innerY += 22
        self.rows = []
        for i in range(5):
            self._add_row(i+1, innerY)
            innerY += 52

        # Панель кнопок справа
        self.pnlButtons = Panel()
        self.pnlButtons.Location = Point(780, 90)
        self.pnlButtons.Size = Size(110, 310)
        self.pnlButtons.BackColor = Color.Transparent
        self.Controls.Add(self.pnlButtons)

        btnY = 10
        self.btnSelect5 = Button()
        self.btnSelect5.Text = u'Выбрать 5 шт'
        self.btnSelect5.Location = Point(0, btnY)
        self.btnSelect5.Size = Size(105, 35)
        self.btnSelect5.FlatStyle = FlatStyle.Flat
        self.btnSelect5.BackColor = COLOR_BTN_SECONDARY
        self.btnSelect5.ForeColor = Color.White
        self.btnSelect5.FlatAppearance.BorderSize = 0
        self.btnSelect5.Click += self._on_select_5
        self.pnlButtons.Controls.Add(self.btnSelect5)

        # Нижняя панель
        self.pnlBottom = Panel()
        self.pnlBottom.Location = Point(0, 410)
        self.pnlBottom.Size = Size(900, 110)
        self.pnlBottom.BackColor = COLOR_HEADER_BG
        self.Controls.Add(self.pnlBottom)

        # Label для отображения количества
        self.lblCount = Label()
        self.lblCount.Text = u''
        self.lblCount.Location = Point(15, 15)
        self.lblCount.Size = Size(500, 25)
        self.lblCount.Font = Font(self.Font.FontFamily, 10, FontStyle.Regular)
        self.pnlBottom.Controls.Add(self.lblCount)

        self.lblInfo = Label()
        self.lblInfo.Text = u'Выберите параметры и значения для фильтрации элементов на виде.'
        self.lblInfo.Location = Point(15, 45)
        self.lblInfo.Size = Size(600, 20)
        self.lblInfo.ForeColor = Color.Gray
        self.pnlBottom.Controls.Add(self.lblInfo)

        self.btnCancel = Button()
        self.btnCancel.Text = u'Закрыть'
        self.btnCancel.Location = Point(675, 65)
        self.btnCancel.Size = Size(100, 35)
        self.btnCancel.FlatStyle = FlatStyle.Flat
        self.btnCancel.BackColor = COLOR_ACCENT_BLUE
        self.btnCancel.ForeColor = Color.White
        self.btnCancel.FlatAppearance.BorderSize = 0
        self.btnCancel.DialogResult = DialogResult.Cancel
        self.CancelButton = self.btnCancel
        self.pnlBottom.Controls.Add(self.btnCancel)

        self.btnOk = Button()
        self.btnOk.Text = u'Выбрать все'
        self.btnOk.Location = Point(780, 65)
        self.btnOk.Size = Size(105, 35)
        self.btnOk.FlatStyle = FlatStyle.Flat
        self.btnOk.BackColor = Color.FromArgb(76, 175, 80)  # Зелёный
        self.btnOk.ForeColor = Color.White
        self.btnOk.FlatAppearance.BorderSize = 0
        self.btnOk.Click += self._on_ok
        self.AcceptButton = self.btnOk
        self.pnlBottom.Controls.Add(self.btnOk)

        self.include_family = False
        self.param_names = []
        self.param_index = {}
        self.values_cache = {}
        self.values_loaded = set()
        self.last_count = 0  # Последнее посчитанное количество
        self.last_conditions = []  # Последние условия для подсчёта

        self._rebuild_index()
        self._prefill_from_selection()  # Предзаполнение из выделенного элемента

    def _add_row(self, idx, y):
        # Панель строки с чередующимся фоном
        rowPanel = Panel()
        rowPanel.Location = Point(5, y)
        rowPanel.Size = Size(745, 48)
        if idx % 2 == 0:
            rowPanel.BackColor = COLOR_ROW_ALT
        else:
            rowPanel.BackColor = Color.White
        self.grpConditions.Controls.Add(rowPanel)

        lbl = Label()
        lbl.Text = u'%d' % idx
        lbl.Location = Point(10, 14)
        lbl.Size = Size(25, 22)
        lbl.Font = Font(self.Font.FontFamily, 10, FontStyle.Bold)
        lbl.ForeColor = COLOR_ACCENT_BLUE
        lbl.TextAlign = ContentAlignment.MiddleCenter
        rowPanel.Controls.Add(lbl)

        cmbP = ComboBox()
        cmbP.DropDownStyle = ComboBoxStyle.DropDown
        cmbP.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        cmbP.AutoCompleteSource = AutoCompleteSource.ListItems
        cmbP.Location = Point(45, 12)
        cmbP.Size = Size(330, 24)
        cmbP.Font = Font(self.Font.FontFamily, 9, FontStyle.Regular)
        cmbP.SelectedIndexChanged += self._on_param_changed
        cmbP.Tag = idx-1
        rowPanel.Controls.Add(cmbP)

        cmbO = ComboBox()
        for o in OPS:
            cmbO.Items.Add(o)
        cmbO.SelectedIndex = 0
        cmbO.DropDownStyle = ComboBoxStyle.DropDownList
        cmbO.Location = Point(385, 12)
        cmbO.Size = Size(110, 24)
        cmbO.Font = Font(self.Font.FontFamily, 9, FontStyle.Regular)
        cmbO.SelectedIndexChanged += self._on_op_changed
        cmbO.Tag = idx-1
        rowPanel.Controls.Add(cmbO)

        cmbV = ComboBox()
        cmbV.DropDownStyle = ComboBoxStyle.DropDownList
        cmbV.Location = Point(505, 12)
        cmbV.Size = Size(230, 24)
        cmbV.Font = Font(self.Font.FontFamily, 9, FontStyle.Regular)
        cmbV.DropDown += self._on_value_dropdown
        cmbV.SelectedIndexChanged += self._on_value_changed
        cmbV.Tag = idx-1
        rowPanel.Controls.Add(cmbV)

        self.rows.append((cmbP, cmbO, cmbV))

    def _on_scope_changed(self, sender, args):
        self._rebuild_index()

    def _on_load_family(self, sender, args):
        self.include_family = True
        self.btnLoadFamily.Enabled = False
        self._rebuild_index()

    def _rebuild_index(self):
        self.values_cache.clear()
        self.values_loaded.clear()
        only_vis = bool(self.cbOnlyVis.Checked)
        include_types = bool(self.cbIncludeTypes.Checked)
        names, pindex = build_param_index(only_vis, include_types, self.include_family, sample_per_cat=20)
        self.param_names = names
        self.param_index = pindex

        for cmbP, cmbO, cmbV in self.rows:
            cmbP.Items.Clear()
            for n in self.param_names:
                cmbP.Items.Add(n)
            cmbP.SelectedIndex = -1
            cmbV.Items.Clear()

    def _prefill_from_selection(self):
        """Предзаполнение фильтра на основе выделенного элемента"""
        global _preselected_element
        if _preselected_element is None:
            return
        
        el = _preselected_element
        
        # Строка 1: Категория
        cat_name = get_element_category_name(el)
        if cat_name:
            cat_param_name = u'Категория'
            if cat_param_name in self.param_names:
                cmbP1, cmbO1, cmbV1 = self.rows[0]
                cmbP1.Text = cat_param_name
                # Загружаем значения для этого параметра
                self._load_values_for_row(0)
                # Устанавливаем значение категории
                for i in range(cmbV1.Items.Count):
                    if _to_unicode(cmbV1.Items[i]) == cat_name:
                        cmbV1.SelectedIndex = i
                        break
        
        # Строка 2: Семейство и типоразмер
        fam_type = get_element_family_and_type(el)
        if fam_type:
            fam_param_name = u'Семейство и типоразмер'
            if fam_param_name in self.param_names:
                cmbP2, cmbO2, cmbV2 = self.rows[1]
                cmbP2.Text = fam_param_name
                # Загружаем значения для этого параметра
                self._load_values_for_row(1)
                # Устанавливаем значение семейства и типоразмера
                for i in range(cmbV2.Items.Count):
                    if _to_unicode(cmbV2.Items[i]) == fam_type:
                        cmbV2.SelectedIndex = i
                        break
        
        # Обновляем счётчик
        self._update_count()

    def _load_values_for_row(self, row):
        """Загрузить значения для параметра в строке"""
        cmbP, cmbO, cmbV = self.rows[row]
        pname = _to_unicode(cmbP.Text)
        if not pname:
            return
        if pname in self.values_loaded:
            return
        only_vis = bool(self.cbOnlyVis.Checked)
        include_types = bool(self.cbIncludeTypes.Checked)
        storage = self.param_index.get(pname, None)
        if storage is None:
            return
        disp_list, mapping, has_empty = collect_values_for_param(
            pname, storage, only_vis, include_types, self.include_family
        )
        self.values_cache[pname] = {'disp_list': disp_list, 'map': mapping, 'has_empty': has_empty}
        self.values_loaded.add(pname)
        cmbV.Items.Clear()
        for v in disp_list:
            cmbV.Items.Add(v)

    def _on_param_changed(self, sender, args):
        row = int(sender.Tag)
        cmbP, cmbO, cmbV = self.rows[row]
        cmbV.Items.Clear()
        pname = _to_unicode(cmbP.Text)
        if pname and pname in self.values_cache:
            for v in self.values_cache[pname]['disp_list']:
                cmbV.Items.Add(v)
        self._sync_value_enabled(row)

    def _on_value_dropdown(self, sender, args):
        row = int(sender.Tag)
        cmbP, cmbO, cmbV = self.rows[row]
        pname = _to_unicode(cmbP.Text)
        if not pname:
            return
        if pname in self.values_loaded:
            return
        only_vis = bool(self.cbOnlyVis.Checked)
        include_types = bool(self.cbIncludeTypes.Checked)
        storage = self.param_index.get(pname, None)
        if storage is None:
            return
        disp_list, mapping, has_empty = collect_values_for_param(
            pname, storage, only_vis, include_types, self.include_family
        )
        self.values_cache[pname] = {'disp_list': disp_list, 'map': mapping, 'has_empty': has_empty}
        self.values_loaded.add(pname)
        cmbV.Items.Clear()
        for v in disp_list:
            cmbV.Items.Add(v)

    def _on_op_changed(self, sender, args):
        row = int(sender.Tag)
        self._sync_value_enabled(row)
        self._update_count()

    def _on_value_changed(self, sender, args):
        """Callback при изменении значения - автоподсчёт"""
        self._update_count()

    def _update_count(self):
        """Автоматический подсчёт элементов при изменении условий"""
        conds = self._get_conditions()
        if not conds:
            self.lblCount.Text = u''
            return
        
        only_vis = bool(self.cbOnlyVis.Checked)
        lookup_in_type = True
        use_or = (_to_unicode(self.cmbLogic.Text) == u'ИЛИ')
        
        count = self._count_matches(conds, only_vis, lookup_in_type, use_or)
        self.last_count = count
        self.last_conditions = conds
        
        if count == 0:
            self.lblCount.Text = u'Элементов не найдено'
            self.lblCount.ForeColor = Color.Gray
        else:
            self.lblCount.Text = u'Будет выбрано элементов: {0}'.format(count)
            self.lblCount.ForeColor = COLOR_ACCENT_BLUE

    def _sync_value_enabled(self, row):
        cmbP, cmbO, cmbV = self.rows[row]
        op = _to_unicode(cmbO.Text)
        need_val = op not in (u'пусто', u'не пусто')
        cmbV.Enabled = need_val

    def _get_conditions(self):
        """Получить список условий из формы"""
        conds = []
        for cmbP, cmbO, cmbV in self.rows:
            pname = _to_unicode(cmbP.Text).strip()
            op = _to_unicode(cmbO.Text).strip()
            if not pname:
                continue
            if op in (u'пусто', u'не пусто'):
                raw = None
            else:
                disp = _to_unicode(cmbV.Text)
                if disp == u'' or disp is None:
                    continue
                cache = self.values_cache.get(pname)
                if cache is None:
                    storage = self.param_index.get(pname, None)
                    only_vis = bool(self.cbOnlyVis.Checked)
                    include_types = bool(self.cbIncludeTypes.Checked)
                    disp_list, mapping, has_empty = collect_values_for_param(
                        pname, storage, only_vis, include_types, self.include_family
                    )
                    cache = {'disp_list': disp_list, 'map': mapping, 'has_empty': has_empty}
                    self.values_cache[pname] = cache
                raw = cache['map'].get(disp, disp)
            conds.append((pname, op, raw))
        return conds

    def _count_matches(self, conditions, only_vis, lookup_in_type, use_or):
        """Подсчитать количество подходящих элементов"""
        elems = collect_candidates(only_vis)
        count = 0
        out = script.get_output()
        try:
            pb = out.create_progress_bar(len(elems), title=u'Подсчёт элементов...')
        except:
            pb = None
        try:
            for el in elems:
                if pb: pb.update()
                if _is_camera(el):
                    continue
                flags = []
                for (pname, op, raw) in conditions:
                    ok = match_condition(el, pname, op, raw, lookup_in_type)
                    flags.append(ok)
                    if use_or and ok:
                        break
                    if (not use_or) and (not ok):
                        break
                res = any(flags) if use_or else all(flags)
                if res:
                    count += 1
        finally:
            if pb: pb.close()
        return count

    def _on_select_5(self, sender, args):
        """Обработчик кнопки выбора 5 элементов"""
        conds = self._get_conditions()
        if not conds:
            from pyrevit import forms
            forms.alert(u'Не выбрано ни одного условия.', title=u'Суперфильтр')
            return

        self.values = {
            'conditions': conds,
            'onlyvis': bool(self.cbOnlyVis.Checked),
            'search_types': True,
            'logic': _to_unicode(self.cmbLogic.Text),
            'limit': 5
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_ok(self, sender, args):
        conds = self._get_conditions()
        if not conds:
            from pyrevit import forms
            forms.alert(u'Не выбрано ни одного условия.', title=u'Суперфильтр')
            return

        self.values = {
            'conditions': conds,
            'onlyvis': bool(self.cbOnlyVis.Checked),
            'search_types': True,
            'logic': _to_unicode(self.cmbLogic.Text),
            'limit': None
        }
        self.DialogResult = DialogResult.OK
        self.Close()

# ---------------- Main ----------------
def main():
    ensure_view_supported(active_view)
    frm = FilterForm()
    res = frm.ShowDialog()
    if res != DialogResult.OK or not getattr(frm, 'values', None):
        return
    ui = frm.values

    conditions = ui['conditions']
    only_vis = ui['onlyvis']
    lookup_in_type = ui['search_types']
    use_or = (ui['logic'] == u'ИЛИ')
    limit = ui.get('limit', None)

    elems = collect_candidates(only_vis)
    matched_ids = []

    out = script.get_output()
    try:
        pb = out.create_progress_bar(len(elems), title=u'Фильтрую элементы...')
    except:
        pb = None

    try:
        for el in elems:
            if pb: pb.update()
            if _is_camera(el):
                continue
            flags = []
            for (pname, op, raw) in conditions:
                ok = match_condition(el, pname, op, raw, lookup_in_type)
                flags.append(ok)
                if use_or and ok:
                    break
                if (not use_or) and (not ok):
                    break
            res = any(flags) if use_or else all(flags)
            if res:
                matched_ids.append(el.Id)
                # Если установлен лимит и достигнут - прерываем поиск
                if limit and len(matched_ids) >= limit:
                    break
    finally:
        if pb: pb.close()

    if not matched_ids:
        from pyrevit import forms
        forms.alert(u'Элементы не найдены по выбранным условиям.', title=u'Суперфильтр')
        return

    from System.Collections.Generic import List as CsList
    sel_ids = CsList[ElementId](matched_ids)
    uidoc.Selection.SetElementIds(sel_ids)
    try:
        uidoc.ShowElements(sel_ids)
    except InvalidOperationException:
        pass

    from pyrevit import forms
    msg = u'Найдено и выбрано элементов: {0}'.format(len(matched_ids))
    if limit:
        msg += u' (ограничение: {0})'.format(limit)
    forms.alert(msg, title=u'Суперфильтр', warn_icon=False)

if __name__ == '__main__':
    main()
