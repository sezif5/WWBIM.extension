# -*- coding: utf-8 -*-
# pyRevit button script: Менеджер листов (v23 — редактирование прямо в таблице видимое)

import clr
import re

# --- AddReferences ДО импортов .NET ---
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitAPIUI')

from System import Array, String
import System.Windows.Forms
from System.Windows.Forms import (Form, ComboBox, Label, Button,
                                  ListView, ColumnHeader, AnchorStyles,
                                  ListViewItem, BorderStyle, FormStartPosition,
                                  ComboBoxStyle, MessageBox, MessageBoxButtons, MessageBoxIcon,
                                  NumericUpDown, Keys, Panel, FlatStyle, Cursors,
                                  CheckBox, RadioButton, GroupBox, DialogResult,
                                  ContextMenuStrip, ToolStripMenuItem,
                                  FolderBrowserDialog, ListBox, SelectionMode)
from System.Windows.Forms import View as WinFormsView
from System.Drawing import Size, Point, Rectangle, Color, Font, FontStyle, Drawing2D
from System.IO import Path

from Autodesk.Revit.DB import (FilteredElementCollector, ViewSheet, BuiltInParameter, BuiltInCategory,
                               Transaction, TransactionGroup, IFailuresPreprocessor,
                               FailureProcessingResult, ElementId, StorageType,
                               Viewport, View, ViewSchedule, CopyPasteOptions, ElementTransformUtils,
                               TextNote, AnnotationSymbol, DetailCurve, FilledRegion,
                               ScheduleSheetInstance, ViewDuplicateOption, ViewType,
                               PrintManager, PrintRange, ViewSet, PrintSetup)

# --- Значения по умолчанию ---
DEFAULT_GROUP_PARAM = u"ADSK_Штамп Раздел проекта"
DEFAULT_NUM_PARAM   = u"DVLK_Штамп_Номер листа"
FORMAT_PARAM_CYR    = u"А"    # Кириллическая 'А'
FORMAT_PARAM_LAT    = u"A"    # Латинская 'A'

# --- Доступ к документу ---
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# --- Helpers ---
def get_param_str(elem, name):
    p = elem.LookupParameter(name)
    if p:
        try:
            return (p.AsString() or u"").strip()
        except:
            pass
    return u""

def set_param_str(elem, name, value):
    p = elem.LookupParameter(name)
    if p and not p.IsReadOnly:
        p.Set(value)
        return True
    return False

def set_param_on_titleblock(tb, name, value):
    """Устанавливает параметр на титульнике (экземпляр или тип)"""
    # Сначала пробуем экземпляр
    p = tb.LookupParameter(name)
    if p and not p.IsReadOnly:
        try:
            p.Set(value)
            return True
        except:
            pass
    # Если не получилось, пробуем тип
    try:
        tb_type = doc.GetElement(tb.GetTypeId())
        if tb_type:
            p = tb_type.LookupParameter(name)
            if p and not p.IsReadOnly:
                try:
                    p.Set(value)
                    return True
                except:
                    pass
    except:
        pass
    return False

def get_sheet_number(sheet):
    p = sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER)
    return ((p.AsString() if p else u"") or u"").strip()

def set_sheet_number(sheet, value):
    p = sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER)
    if p and not p.IsReadOnly:
        p.Set(value)
        return True
    return False

def is_placeholder(sheet):
    try:
        return sheet.IsPlaceholder
    except:
        return False

def get_sheets(document):
    return [s for s in FilteredElementCollector(document).OfClass(ViewSheet).ToElements() if not is_placeholder(s)]

def get_titleblocks_on_sheet(sheet):
    col = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType()
    res = []
    for tb in col:
        try:
            if tb.OwnerViewId == sheet.Id:
                res.append(tb)
        except:
            pass
    return res

def get_format_value(sheet):
    tblks = get_titleblocks_on_sheet(sheet)
    if not tblks:
        return u""
    tb = tblks[0]
    p = tb.LookupParameter(FORMAT_PARAM_CYR) or tb.LookupParameter(FORMAT_PARAM_LAT)
    if not p:
        return u""
    try:
        if p.StorageType == StorageType.Integer:
            return unicode(p.AsInteger())
        else:
            return (p.AsString() or u"")
    except:
        return u""

def set_format_value(sheet, value_text):
    tblks = get_titleblocks_on_sheet(sheet)
    if not tblks:
        return False
    ok_any = False
    for tb in tblks:
        p = tb.LookupParameter(FORMAT_PARAM_CYR) or tb.LookupParameter(FORMAT_PARAM_LAT)
        if not p or p.IsReadOnly:
            continue
        try:
            if p.StorageType == StorageType.Integer:
                ival = int(value_text)
                p.Set(ival)
                ok_any = True
            else:
                p.Set(value_text)
                ok_any = True
        except:
            pass
    return ok_any

def get_titleblock_type_name(sheet):
    """Получает имя типа семейства основной надписи"""
    tblks = get_titleblocks_on_sheet(sheet)
    if not tblks:
        return u""
    tb = tblks[0]
    try:
        tb_type = doc.GetElement(tb.GetTypeId())
        if tb_type:
            return tb_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or u""
    except:
        pass
    return u""

def get_all_titleblock_types():
    """Получает список всех типов семейств основных надписей в проекте"""
    types_dict = {}  # name -> ElementId
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()
    for tb_type in collector:
        try:
            name = tb_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if name:
                types_dict[name] = tb_type.Id
        except:
            pass
    return types_dict

def set_titleblock_type(sheet, type_name, types_dict):
    """Изменяет тип семейства основной надписи на листе"""
    if type_name not in types_dict:
        return False
    new_type_id = types_dict[type_name]
    tblks = get_titleblocks_on_sheet(sheet)
    if not tblks:
        return False
    ok_any = False
    for tb in tblks:
        try:
            tb.ChangeTypeId(new_type_id)
            ok_any = True
        except:
            pass
    return ok_any

_num_re = re.compile(r"(\d+)")
def natural_key(s):
    if s is None:
        s = u""
    parts = _num_re.split(s)
    key = []
    for i, part in enumerate(parts):
        if i % 2 == 0:
            key.append(part)
        else:
            try:
                key.append(int(part))
            except:
                key.append(part)
    return tuple(key)

# --- Failures suppressor (минимальный) ---
class SuppressErrors(IFailuresPreprocessor):
    def PreprocessFailures(self, fa):
        return FailureProcessingResult.Continue

def with_suppress(t):
    fho = t.GetFailureHandlingOptions()
    fho = fho.SetFailuresPreprocessor(SuppressErrors())
    t.SetFailureHandlingOptions(fho)

# --- Модель строки ---
class SheetRow(object):
    __slots__ = ("sheet_id", "sheetnum", "name")
    def __init__(self, sheet):
        self.sheet_id = sheet.Id.IntegerValue
        self.sheetnum = get_sheet_number(sheet) or u""
        self.name = sheet.Name or u""

# --- Сбор имён параметров ---
def collect_string_param_names(sheets):
    names_all = set()
    names_writable = set()
    for s in sheets:
        try:
            for p in s.Parameters:
                d = p.Definition
                if d is None:
                    continue
                nm = (d.Name or u"").strip()
                if not nm:
                    continue
                try:
                    _ = p.AsString()
                    names_all.add(nm)
                    if not p.IsReadOnly:
                        names_writable.add(nm)
                except:
                    pass
        except:
            pass
    return sorted(names_all, key=natural_key), sorted(names_writable, key=natural_key)


# --- Диалог выбора содержимого для копирования ---
class CopyContentDialog(Form):
    """Диалог выбора что копировать с листа"""
    
    # Константы для режимов легенд/спецификаций
    MODE_NOTHING = 0
    MODE_COPY = 1
    MODE_PLACE_SAME = 2
    
    def __init__(self, src_sheet, legends_on_sheet, schedules_on_sheet):
        self.src_sheet = src_sheet
        self.legends_on_sheet = legends_on_sheet  # список (viewport, view_name)
        self.schedules_on_sheet = schedules_on_sheet  # список (schedule_instance, schedule_name)
        
        self.result_ok = False
        self.copy_viewports = False
        self.copy_text_notes = False
        self.copy_annotations = False
        self.copy_detail_lines = False
        self.legends_mode = {}  # {view_name: MODE_*}
        self.schedules_mode = {}  # {schedule_name: MODE_*}
        
        self._init_ui()
    
    def _init_ui(self):
        self.Text = u"Копировать с содержимым"
        self.Size = Size(500, 500)
        self.MinimumSize = Size(450, 400)
        self.StartPosition = FormStartPosition.CenterParent
        self.BackColor = Color.FromArgb(240, 244, 248)
        try:
            self.Font = Font("Segoe UI", 9.0)
        except:
            pass
        
        y = 15
        
        # === Группа: Элементы для копирования ===
        grp_elements = GroupBox()
        grp_elements.Text = u"Элементы для копирования"
        grp_elements.Location = Point(15, y)
        grp_elements.Size = Size(455, 120)
        grp_elements.ForeColor = Color.FromArgb(60, 64, 67)
        self.Controls.Add(grp_elements)
        
        self.chk_viewports = CheckBox()
        self.chk_viewports.Text = u"Видовые экраны (создать копии видов)"
        self.chk_viewports.Location = Point(15, 25)
        self.chk_viewports.Size = Size(420, 22)
        self.chk_viewports.Checked = False
        grp_elements.Controls.Add(self.chk_viewports)
        
        self.chk_text_notes = CheckBox()
        self.chk_text_notes.Text = u"Текстовые примечания"
        self.chk_text_notes.Location = Point(15, 50)
        self.chk_text_notes.Size = Size(420, 22)
        self.chk_text_notes.Checked = True
        grp_elements.Controls.Add(self.chk_text_notes)
        
        self.chk_annotations = CheckBox()
        self.chk_annotations.Text = u"Аннотации (марки, обозначения)"
        self.chk_annotations.Location = Point(15, 75)
        self.chk_annotations.Size = Size(420, 22)
        self.chk_annotations.Checked = True
        grp_elements.Controls.Add(self.chk_annotations)
        
        self.chk_detail_lines = CheckBox()
        self.chk_detail_lines.Text = u"Линии детализации"
        self.chk_detail_lines.Location = Point(240, 75)
        self.chk_detail_lines.Size = Size(200, 22)
        self.chk_detail_lines.Checked = True
        grp_elements.Controls.Add(self.chk_detail_lines)
        
        y += 135
        
        # === Группа: Легенды ===
        self.legend_radios = {}
        if self.legends_on_sheet:
            grp_legends = GroupBox()
            grp_legends.Text = u"Легенды"
            grp_legends.Location = Point(15, y)
            legends_height = 30 + len(self.legends_on_sheet) * 30
            grp_legends.Size = Size(455, min(legends_height, 180))
            grp_legends.ForeColor = Color.FromArgb(60, 64, 67)
            self.Controls.Add(grp_legends)
            
            ly = 22
            for vp, view_name in self.legends_on_sheet:
                # Создаём Panel для группировки радиокнопок этой строки
                row_panel = Panel()
                row_panel.Location = Point(0, ly)
                row_panel.Size = Size(450, 26)
                grp_legends.Controls.Add(row_panel)
                
                lbl = Label()
                lbl.Text = view_name[:25] + u"..." if len(view_name) > 25 else view_name
                lbl.Location = Point(15, 3)
                lbl.Size = Size(200, 20)
                lbl.ForeColor = Color.FromArgb(40, 44, 47)
                row_panel.Controls.Add(lbl)
                
                rb_nothing = RadioButton()
                rb_nothing.Text = u"Ничего"
                rb_nothing.Location = Point(220, 1)
                rb_nothing.Size = Size(70, 22)
                rb_nothing.Checked = True
                row_panel.Controls.Add(rb_nothing)
                
                rb_same = RadioButton()
                rb_same.Text = u"Ту же"
                rb_same.Location = Point(295, 1)
                rb_same.Size = Size(70, 22)
                row_panel.Controls.Add(rb_same)
                
                rb_copy = RadioButton()
                rb_copy.Text = u"Копия"
                rb_copy.Location = Point(365, 1)
                rb_copy.Size = Size(70, 22)
                row_panel.Controls.Add(rb_copy)
                
                self.legend_radios[view_name] = (rb_nothing, rb_same, rb_copy)
                ly += 26
            
            y += grp_legends.Height + 15
        
        # === Группа: Спецификации ===
        self.schedule_radios = {}
        if self.schedules_on_sheet:
            grp_schedules = GroupBox()
            grp_schedules.Text = u"Спецификации"
            grp_schedules.Location = Point(15, y)
            schedules_height = 30 + len(self.schedules_on_sheet) * 30
            grp_schedules.Size = Size(455, min(schedules_height, 180))
            grp_schedules.ForeColor = Color.FromArgb(60, 64, 67)
            self.Controls.Add(grp_schedules)
            
            sy = 22
            for sched_inst, sched_name in self.schedules_on_sheet:
                # Создаём Panel для группировки радиокнопок этой строки
                row_panel = Panel()
                row_panel.Location = Point(0, sy)
                row_panel.Size = Size(450, 26)
                grp_schedules.Controls.Add(row_panel)
                
                lbl = Label()
                lbl.Text = sched_name[:25] + u"..." if len(sched_name) > 25 else sched_name
                lbl.Location = Point(15, 3)
                lbl.Size = Size(200, 20)
                lbl.ForeColor = Color.FromArgb(40, 44, 47)
                row_panel.Controls.Add(lbl)
                
                rb_nothing = RadioButton()
                rb_nothing.Text = u"Ничего"
                rb_nothing.Location = Point(220, 1)
                rb_nothing.Size = Size(70, 22)
                rb_nothing.Checked = True
                row_panel.Controls.Add(rb_nothing)
                
                rb_same = RadioButton()
                rb_same.Text = u"Ту же"
                rb_same.Location = Point(295, 1)
                rb_same.Size = Size(70, 22)
                row_panel.Controls.Add(rb_same)
                
                rb_copy = RadioButton()
                rb_copy.Text = u"Копия"
                rb_copy.Location = Point(365, 1)
                rb_copy.Size = Size(70, 22)
                row_panel.Controls.Add(rb_copy)
                
                self.schedule_radios[sched_name] = (rb_nothing, rb_same, rb_copy)
                sy += 26
            
            y += grp_schedules.Height + 15
        
        # Увеличим форму если много контента
        if y > 350:
            self.Size = Size(500, y + 100)
        
        # === Кнопки ===
        self.btn_ok = Button()
        self.btn_ok.Text = u"Копировать"
        self.btn_ok.Size = Size(110, 35)
        self.btn_ok.Location = Point(240, y + 10)
        self.btn_ok.BackColor = Color.FromArgb(52, 168, 83)
        self.btn_ok.ForeColor = Color.White
        self.btn_ok.FlatStyle = FlatStyle.Flat
        self.btn_ok.FlatAppearance.BorderSize = 0
        self.btn_ok.Click += self.on_ok
        self.Controls.Add(self.btn_ok)
        
        self.btn_cancel = Button()
        self.btn_cancel.Text = u"Отмена"
        self.btn_cancel.Size = Size(100, 35)
        self.btn_cancel.Location = Point(360, y + 10)
        self.btn_cancel.BackColor = Color.FromArgb(95, 99, 104)
        self.btn_cancel.ForeColor = Color.White
        self.btn_cancel.FlatStyle = FlatStyle.Flat
        self.btn_cancel.FlatAppearance.BorderSize = 0
        self.btn_cancel.Click += self.on_cancel
        self.Controls.Add(self.btn_cancel)
    
    def on_ok(self, sender, e):
        self.result_ok = True
        self.copy_viewports = self.chk_viewports.Checked
        self.copy_text_notes = self.chk_text_notes.Checked
        self.copy_annotations = self.chk_annotations.Checked
        self.copy_detail_lines = self.chk_detail_lines.Checked
        
        # Собираем режимы для легенд
        for view_name, (rb_nothing, rb_same, rb_copy) in self.legend_radios.items():
            if rb_copy.Checked:
                self.legends_mode[view_name] = self.MODE_COPY
            elif rb_same.Checked:
                self.legends_mode[view_name] = self.MODE_PLACE_SAME
            else:
                self.legends_mode[view_name] = self.MODE_NOTHING
        
        # Собираем режимы для спецификаций
        for sched_name, (rb_nothing, rb_same, rb_copy) in self.schedule_radios.items():
            if rb_copy.Checked:
                self.schedules_mode[sched_name] = self.MODE_COPY
            elif rb_same.Checked:
                self.schedules_mode[sched_name] = self.MODE_PLACE_SAME
            else:
                self.schedules_mode[sched_name] = self.MODE_NOTHING
        
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_cancel(self, sender, e):
        self.result_ok = False
        self.DialogResult = DialogResult.Cancel
        self.Close()


# --- Диалог выбора папки для PDF ---
class PDFExportDialog(Form):
    """Диалог настроек экспорта в PDF"""
    
    def __init__(self, album_name, sheets_count):
        self.album_name = album_name
        self.sheets_count = sheets_count
        self.result_ok = False
        self.selected_folder = None
        self.file_name = None
        self.combine_pdf = True
        
        self._init_ui()
    
    def _init_ui(self):
        self.Text = u"Экспорт в PDF"
        self.Size = Size(500, 280)
        self.MinimumSize = Size(450, 260)
        self.StartPosition = FormStartPosition.CenterParent
        self.BackColor = Color.FromArgb(240, 244, 248)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        try:
            self.Font = Font("Segoe UI", 9.0)
        except:
            pass
        
        y = 20
        
        # Информация об альбоме
        self.lbl_info = Label()
        self.lbl_info.Text = u"Альбом: {0}  |  Листов: {1}".format(self.album_name, self.sheets_count)
        self.lbl_info.Location = Point(20, y)
        self.lbl_info.Size = Size(440, 25)
        self.lbl_info.ForeColor = Color.FromArgb(60, 64, 67)
        try:
            self.lbl_info.Font = Font("Segoe UI", 10.0, FontStyle.Bold)
        except:
            pass
        self.Controls.Add(self.lbl_info)
        
        y += 40
        
        # Выбор папки
        self.lbl_folder = Label()
        self.lbl_folder.Text = u"Папка:"
        self.lbl_folder.Location = Point(20, y)
        self.lbl_folder.Size = Size(100, 22)
        self.lbl_folder.ForeColor = Color.FromArgb(60, 64, 67)
        self.Controls.Add(self.lbl_folder)
        
        self.txt_folder = System.Windows.Forms.TextBox()
        self.txt_folder.Location = Point(130, y - 3)
        self.txt_folder.Size = Size(260, 25)
        self.txt_folder.BackColor = Color.White
        self.txt_folder.ForeColor = Color.FromArgb(60, 64, 67)
        self.txt_folder.ReadOnly = True
        self.Controls.Add(self.txt_folder)
        
        self.btn_browse = Button()
        self.btn_browse.Text = u"..."
        self.btn_browse.Location = Point(395, y - 4)
        self.btn_browse.Size = Size(65, 27)
        self.btn_browse.BackColor = Color.FromArgb(95, 99, 104)
        self.btn_browse.ForeColor = Color.White
        self.btn_browse.FlatStyle = FlatStyle.Flat
        self.btn_browse.FlatAppearance.BorderSize = 0
        self.btn_browse.Click += self.on_browse
        self.Controls.Add(self.btn_browse)
        
        y += 40
        
        # Имя файла
        self.lbl_filename = Label()
        self.lbl_filename.Text = u"Имя файла:"
        self.lbl_filename.Location = Point(20, y)
        self.lbl_filename.Size = Size(100, 22)
        self.lbl_filename.ForeColor = Color.FromArgb(60, 64, 67)
        self.Controls.Add(self.lbl_filename)
        
        self.txt_filename = System.Windows.Forms.TextBox()
        self.txt_filename.Location = Point(130, y - 3)
        self.txt_filename.Size = Size(330, 25)
        self.txt_filename.BackColor = Color.White
        self.txt_filename.ForeColor = Color.FromArgb(60, 64, 67)
        # Предлагаем имя по умолчанию
        self.txt_filename.Text = self.album_name.replace(u"/", u"-").replace(u"\\", u"-").replace(u":", u"-")
        self.Controls.Add(self.txt_filename)
        
        y += 40
        
        # Чекбокс объединения
        self.chk_combine = System.Windows.Forms.CheckBox()
        self.chk_combine.Text = u"Объединить все листы в один PDF"
        self.chk_combine.Location = Point(130, y)
        self.chk_combine.Size = Size(300, 25)
        self.chk_combine.Checked = True
        self.chk_combine.ForeColor = Color.FromArgb(60, 64, 67)
        self.Controls.Add(self.chk_combine)
        
        y += 45
        
        # Кнопки
        self.btn_ok = Button()
        self.btn_ok.Text = u"Экспорт PDF"
        self.btn_ok.Size = Size(120, 38)
        self.btn_ok.Location = Point(240, y)
        self.btn_ok.BackColor = Color.FromArgb(234, 67, 53)
        self.btn_ok.ForeColor = Color.White
        self.btn_ok.FlatStyle = FlatStyle.Flat
        self.btn_ok.FlatAppearance.BorderSize = 0
        self.btn_ok.Click += self.on_ok
        self.Controls.Add(self.btn_ok)
        
        self.btn_cancel = Button()
        self.btn_cancel.Text = u"Отмена"
        self.btn_cancel.Size = Size(100, 38)
        self.btn_cancel.Location = Point(365, y)
        self.btn_cancel.BackColor = Color.FromArgb(95, 99, 104)
        self.btn_cancel.ForeColor = Color.White
        self.btn_cancel.FlatStyle = FlatStyle.Flat
        self.btn_cancel.FlatAppearance.BorderSize = 0
        self.btn_cancel.Click += self.on_cancel
        self.Controls.Add(self.btn_cancel)
    
    def on_browse(self, sender, e):
        """Выбор папки для сохранения"""
        dlg = FolderBrowserDialog()
        dlg.Description = u"Выберите папку для сохранения PDF"
        dlg.ShowNewFolderButton = True
        
        if dlg.ShowDialog() == DialogResult.OK:
            self.txt_folder.Text = dlg.SelectedPath
    
    def on_ok(self, sender, e):
        # Валидация
        if not self.txt_folder.Text or not self.txt_folder.Text.strip():
            MessageBox.Show(u"Выберите папку для сохранения PDF.", u"PDF Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        if not self.txt_filename.Text or not self.txt_filename.Text.strip():
            MessageBox.Show(u"Введите имя файла.", u"PDF Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        self.result_ok = True
        self.selected_folder = self.txt_folder.Text.strip()
        self.combine_pdf = self.chk_combine.Checked
        
        # Очищаем имя файла от недопустимых символов
        filename = self.txt_filename.Text.strip()
        for c in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']:
            filename = filename.replace(c, '-')
        self.file_name = filename
        
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_cancel(self, sender, e):
        self.result_ok = False
        self.DialogResult = DialogResult.Cancel
        self.Close()


# --- Диалог выбора папки и пресета для DWG ---
class DWGExportDialog(Form):
    """Диалог настроек экспорта в DWG"""
    
    def __init__(self, album_name, sheets_count):
        self.album_name = album_name
        self.sheets_count = sheets_count
        self.result_ok = False
        self.selected_folder = None
        self.selected_setup = None
        
        self._init_ui()
    
    def _init_ui(self):
        self.Text = u"Экспорт в DWG"
        self.Size = Size(500, 250)
        self.MinimumSize = Size(450, 230)
        self.StartPosition = FormStartPosition.CenterParent
        self.BackColor = Color.FromArgb(240, 244, 248)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        try:
            self.Font = Font("Segoe UI", 9.0)
        except:
            pass
        
        y = 20
        
        # Информация об альбоме
        self.lbl_info = Label()
        self.lbl_info.Text = u"Альбом: {0}  |  Листов: {1}".format(self.album_name, self.sheets_count)
        self.lbl_info.Location = Point(20, y)
        self.lbl_info.Size = Size(440, 25)
        self.lbl_info.ForeColor = Color.FromArgb(60, 64, 67)
        try:
            self.lbl_info.Font = Font("Segoe UI", 10.0, FontStyle.Bold)
        except:
            pass
        self.Controls.Add(self.lbl_info)
        
        y += 40
        
        # Выбор пресета экспорта
        self.lbl_setup = Label()
        self.lbl_setup.Text = u"Пресет:"
        self.lbl_setup.Location = Point(20, y)
        self.lbl_setup.Size = Size(100, 22)
        self.lbl_setup.ForeColor = Color.FromArgb(60, 64, 67)
        self.Controls.Add(self.lbl_setup)
        
        self.cmb_setup = ComboBox()
        self.cmb_setup.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_setup.Location = Point(130, y - 3)
        self.cmb_setup.Size = Size(330, 25)
        self.cmb_setup.FlatStyle = FlatStyle.Flat
        self.cmb_setup.BackColor = Color.White
        self.cmb_setup.ForeColor = Color.FromArgb(60, 64, 67)
        self.Controls.Add(self.cmb_setup)
        
        # Заполняем список пресетов
        self._fill_export_setups()
        
        y += 40
        
        # Выбор папки
        self.lbl_folder = Label()
        self.lbl_folder.Text = u"Папка:"
        self.lbl_folder.Location = Point(20, y)
        self.lbl_folder.Size = Size(100, 22)
        self.lbl_folder.ForeColor = Color.FromArgb(60, 64, 67)
        self.Controls.Add(self.lbl_folder)
        
        self.txt_folder = System.Windows.Forms.TextBox()
        self.txt_folder.Location = Point(130, y - 3)
        self.txt_folder.Size = Size(260, 25)
        self.txt_folder.BackColor = Color.White
        self.txt_folder.ForeColor = Color.FromArgb(60, 64, 67)
        self.txt_folder.ReadOnly = True
        self.Controls.Add(self.txt_folder)
        
        self.btn_browse = Button()
        self.btn_browse.Text = u"..."
        self.btn_browse.Location = Point(395, y - 4)
        self.btn_browse.Size = Size(65, 27)
        self.btn_browse.BackColor = Color.FromArgb(95, 99, 104)
        self.btn_browse.ForeColor = Color.White
        self.btn_browse.FlatStyle = FlatStyle.Flat
        self.btn_browse.FlatAppearance.BorderSize = 0
        self.btn_browse.Click += self.on_browse
        self.Controls.Add(self.btn_browse)
        
        y += 50
        
        # Кнопки
        self.btn_ok = Button()
        self.btn_ok.Text = u"Экспорт DWG"
        self.btn_ok.Size = Size(120, 38)
        self.btn_ok.Location = Point(240, y)
        self.btn_ok.BackColor = Color.FromArgb(128, 128, 0)  # Оливковый
        self.btn_ok.ForeColor = Color.White
        self.btn_ok.FlatStyle = FlatStyle.Flat
        self.btn_ok.FlatAppearance.BorderSize = 0
        self.btn_ok.Click += self.on_ok
        self.Controls.Add(self.btn_ok)
        
        self.btn_cancel = Button()
        self.btn_cancel.Text = u"Отмена"
        self.btn_cancel.Size = Size(100, 38)
        self.btn_cancel.Location = Point(365, y)
        self.btn_cancel.BackColor = Color.FromArgb(95, 99, 104)
        self.btn_cancel.ForeColor = Color.White
        self.btn_cancel.FlatStyle = FlatStyle.Flat
        self.btn_cancel.FlatAppearance.BorderSize = 0
        self.btn_cancel.Click += self.on_cancel
        self.Controls.Add(self.btn_cancel)
    
    def _fill_export_setups(self):
        """Заполнение списка пресетов DWG экспорта"""
        try:
            from Autodesk.Revit.DB import ExportDWGSettings
            
            # Получаем все настройки экспорта DWG из документа
            collector = FilteredElementCollector(doc).OfClass(ExportDWGSettings)
            setups = list(collector)
            
            if setups:
                for setup in setups:
                    self.cmb_setup.Items.Add(setup.Name)
                self.cmb_setup.SelectedIndex = 0
            else:
                # Если нет сохранённых пресетов, добавляем дефолтный
                self.cmb_setup.Items.Add(u"<По умолчанию>")
                self.cmb_setup.SelectedIndex = 0
        except Exception as ex:
            self.cmb_setup.Items.Add(u"<По умолчанию>")
            self.cmb_setup.SelectedIndex = 0
    
    def on_browse(self, sender, e):
        """Выбор папки для сохранения"""
        dlg = FolderBrowserDialog()
        dlg.Description = u"Выберите папку для сохранения DWG"
        dlg.ShowNewFolderButton = True
        
        if dlg.ShowDialog() == DialogResult.OK:
            self.txt_folder.Text = dlg.SelectedPath
    
    def on_ok(self, sender, e):
        # Валидация
        if not self.txt_folder.Text or not self.txt_folder.Text.strip():
            MessageBox.Show(u"Выберите папку для сохранения DWG.", u"DWG Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        if self.cmb_setup.SelectedItem is None:
            MessageBox.Show(u"Выберите пресет экспорта.", u"DWG Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        self.result_ok = True
        self.selected_folder = self.txt_folder.Text.strip()
        self.selected_setup = str(self.cmb_setup.SelectedItem)
        
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_cancel(self, sender, e):
        self.result_ok = False
        self.DialogResult = DialogResult.Cancel
        self.Close()


# --- Форма ---
class SheetManagerForm(Form):
    def __init__(self):
        self.Text = u"Менеджер листов | Нумерация"
        self.MinimumSize = Size(980, 640)
        self.Size = Size(1040, 680)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.FromArgb(240, 244, 248)
        
        # Современный шрифт
        try:
            self.Font = Font("Segoe UI", 9.0)
        except:
            pass

        # Контролы параметров
        self.lblSortParam = Label()
        self.lblSortParam.Text = u"Параметр сортировки:"
        self.lblSortParam.AutoSize = True
        self.lblSortParam.Location = Point(12, 14)
        self.lblSortParam.ForeColor = Color.FromArgb(60, 64, 67)
        self.lblSortParam.BackColor = Color.Transparent
        
        self.cmbSortParam = ComboBox()
        self.cmbSortParam.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmbSortParam.Location = Point(160, 10)
        self.cmbSortParam.Width = 220
        self.cmbSortParam.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.cmbSortParam.FlatStyle = FlatStyle.Flat
        self.cmbSortParam.BackColor = Color.White
        self.cmbSortParam.ForeColor = Color.FromArgb(60, 64, 67)

        self.lblGroupVal = Label()
        self.lblGroupVal.Text = u"Альбом:"
        self.lblGroupVal.AutoSize = True
        self.lblGroupVal.Location = Point(390, 14)
        self.lblGroupVal.ForeColor = Color.FromArgb(60, 64, 67)
        self.lblGroupVal.BackColor = Color.Transparent
        
        self.cmbGroupVal = ComboBox()
        self.cmbGroupVal.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmbGroupVal.Location = Point(460, 10)
        self.cmbGroupVal.Width = 220
        self.cmbGroupVal.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.cmbGroupVal.FlatStyle = FlatStyle.Flat
        self.cmbGroupVal.BackColor = Color.White
        self.cmbGroupVal.ForeColor = Color.FromArgb(60, 64, 67)

        self.lblStart = Label()
        self.lblStart.Text = u"Начинать с:"
        self.lblStart.AutoSize = True
        self.lblStart.Location = Point(460, 40)
        self.lblStart.ForeColor = Color.FromArgb(60, 64, 67)
        self.lblStart.BackColor = Color.Transparent
        
        self.numStart = NumericUpDown()
        self.numStart.Minimum = 1
        self.numStart.Maximum = 9999
        self.numStart.Value = 1
        self.numStart.Width = 80
        self.numStart.Location = Point(540, 38)
        self.numStart.BorderStyle = BorderStyle.FixedSingle
        self.numStart.BackColor = Color.White
        self.numStart.ForeColor = Color.FromArgb(60, 64, 67)

        self.lblNumParam = Label()
        self.lblNumParam.Text = u"Параметр нумерации:"
        self.lblNumParam.AutoSize = True
        self.lblNumParam.Location = Point(12, 44)
        self.lblNumParam.ForeColor = Color.FromArgb(60, 64, 67)
        self.lblNumParam.BackColor = Color.Transparent
        
        self.cmbNumParam = ComboBox()
        self.cmbNumParam.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmbNumParam.Location = Point(160, 40)
        self.cmbNumParam.Width = 220
        self.cmbNumParam.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.cmbNumParam.FlatStyle = FlatStyle.Flat
        self.cmbNumParam.BackColor = Color.White
        self.cmbNumParam.ForeColor = Color.FromArgb(60, 64, 67)

        # Таблица
        self.lv = ListView()
        self.lv.View = WinFormsView.Details
        self.lv.FullRowSelect = True
        self.lv.MultiSelect = False
        self.lv.BorderStyle = BorderStyle.FixedSingle
        self.lv.Location = Point(12, 100)
        self.lv.Size = Size(900, 520)
        self.lv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.lv.BackColor = Color.White
        self.lv.ForeColor = Color.FromArgb(60, 64, 67)

        ch1 = ColumnHeader(); ch1.Text = u"Порядок"; ch1.Width = 70
        ch2 = ColumnHeader(); ch2.Text = u"Сист. номер"; ch2.Width = 95
        ch3 = ColumnHeader(); ch3.Text = u"Имя листа"; ch3.Width = 340
        ch4 = ColumnHeader(); ch4.Text = DEFAULT_NUM_PARAM; ch4.Width = 100
        ch5 = ColumnHeader(); ch5.Text = u"Тип надписи"; ch5.Width = 180
        ch6 = ColumnHeader(); ch6.Text = u"Формат (A)"; ch6.Width = 80
        self.col_num = ch4
        self.col_tb_type = ch5
        self.col_fmt = ch6
        self.lv.Columns.AddRange(Array[ColumnHeader]([ch1, ch2, ch3, ch4, ch5, ch6]))

        # Правая колонка кнопок
        self.btnUp = self._create_modern_button(u"▲ Вверх", Size(130, 36), Color.FromArgb(26, 115, 232))
        self.btnDown = self._create_modern_button(u"▼ Вниз", Size(130, 36), Color.FromArgb(26, 115, 232))
        self.btnCopy = self._create_modern_button(u"Копировать", Size(130, 36), Color.FromArgb(26, 115, 232))
        self.btnCopyContent = self._create_modern_button(u"Копировать+", Size(130, 36), Color.FromArgb(66, 133, 244))
        self.btnPDF = self._create_modern_button(u"PDF", Size(130, 36), Color.FromArgb(234, 67, 53))
        self.btnDWG = self._create_modern_button(u"DWG", Size(130, 36), Color.FromArgb(128, 128, 0))  # Оливковый
        self.btnNumerate = self._create_modern_button(u"Нумеровать", Size(130, 40), Color.FromArgb(52, 168, 83))

        for b in [self.btnUp, self.btnDown, self.btnCopy, self.btnCopyContent, self.btnPDF, self.btnDWG]:
            b.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.btnNumerate.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        for c in [self.lblSortParam, self.cmbSortParam, self.lblGroupVal, self.cmbGroupVal, self.lblStart, self.numStart,
                  self.lblNumParam, self.cmbNumParam, self.lv, self.btnUp, self.btnDown, self.btnCopy, self.btnCopyContent, self.btnPDF, self.btnDWG, self.btnNumerate]:
            self.Controls.Add(c)
        
        # Контекстное меню для таблицы
        self.context_menu = ContextMenuStrip()
        self.menu_delete = ToolStripMenuItem(u"Удалить лист")
        self.menu_delete.Click += self.on_delete_sheet
        self.context_menu.Items.Add(self.menu_delete)
        self.lv.ContextMenuStrip = self.context_menu

        # Редактор: делаем дочерним контролом списка (чтобы точно был над сабайтем)
        from System.Windows.Forms import TextBox as WinTextBox
        self.editBox = WinTextBox()
        self.editBox.Visible = False
        self.editBox.BorderStyle = BorderStyle.FixedSingle
        self.editBox.BackColor = Color.White
        self.editBox.ForeColor = Color.Black
        self.editBox.Font = self.lv.Font
        self.editBox.KeyDown += self._editbox_keydown
        self.editBox.Leave += self._editbox_leave
        self.lv.Controls.Add(self.editBox)  # ключевое отличие
        self._edit_ctx = None

        # Данные
        self.sheets = [SheetRow(s) for s in get_sheets(doc)]
        self.sheet_map = {r.sheet_id: doc.GetElement(ElementId(r.sheet_id)) for r in self.sheets}

        names_all, names_writable = collect_string_param_names(self.sheet_map.values())
        if DEFAULT_GROUP_PARAM not in names_all:
            names_all.insert(0, DEFAULT_GROUP_PARAM)
        if DEFAULT_NUM_PARAM not in names_writable:
            names_writable.insert(0, DEFAULT_NUM_PARAM)
        for n in names_all:
            self.cmbSortParam.Items.Add(n)
        for n in names_writable:
            self.cmbNumParam.Items.Add(n)
        try:
            self.cmbSortParam.SelectedIndex = max(0, list(self.cmbSortParam.Items).IndexOf(DEFAULT_GROUP_PARAM))
        except:
            self.cmbSortParam.SelectedIndex = 0
        try:
            self.cmbNumParam.SelectedIndex = max(0, list(self.cmbNumParam.Items).IndexOf(DEFAULT_NUM_PARAM))
        except:
            self.cmbNumParam.SelectedIndex = 0

        # События
        self.cmbSortParam.SelectedIndexChanged += self.on_sort_param_changed
        self.cmbGroupVal.SelectedIndexChanged  += self.on_group_value_changed
        self.cmbNumParam.SelectedIndexChanged  += self.on_num_param_changed
        self.lv.MouseDoubleClick += self.on_lv_double_click
        self.lv.ColumnWidthChanged += self._hide_editor_event
        self.lv.SizeChanged += self._hide_editor_event
        self.lv.MouseWheel += self._hide_editor_event
        self.btnUp.Click += self.on_move_up
        self.btnDown.Click += self.on_move_down
        self.btnCopy.Click += self.on_copy_below
        self.btnCopyContent.Click += self.on_copy_with_content
        self.btnPDF.Click += self.on_pdf_export
        self.btnDWG.Click += self.on_dwg_export
        self.btnNumerate.Click += self.on_numerate
        self.Resize += self.on_resize

        # Инициал
        self.order_ids = []
        self.current_rows = []
        self.rebuild_group_values()
        self.reload_rows(set_initial_order=True)
        self._layout_controls()

    def _create_modern_button(self, text, size, base_color):
        """Создание кнопки в современном стиле"""
        btn = Button()
        btn.Text = text
        btn.Size = size
        btn.FlatStyle = FlatStyle.Flat
        btn.BackColor = base_color
        btn.ForeColor = Color.White
        btn.FlatAppearance.BorderSize = 0
        btn.FlatAppearance.MouseOverBackColor = self._lighten_color(base_color, 20)
        btn.FlatAppearance.MouseDownBackColor = self._darken_color(base_color, 20)
        btn.Cursor = Cursors.Hand
        try:
            btn.Font = Font("Segoe UI", 9.0, FontStyle.Regular)
        except:
            pass
        return btn

    def _lighten_color(self, color, amount):
        """Осветление цвета"""
        r = min(255, color.R + amount)
        g = min(255, color.G + amount)
        b = min(255, color.B + amount)
        return Color.FromArgb(r, g, b)

    def _darken_color(self, color, amount):
        """Затемнение цвета"""
        r = max(0, color.R - amount)
        g = max(0, color.G - amount)
        b = max(0, color.B - amount)
        return Color.FromArgb(r, g, b)

    # Размещение
    def _layout_controls(self):
        right_margin = 20
        bottom_margin = 20
        gap = 10
        client_w = self.ClientSize.Width
        client_h = self.ClientSize.Height
        x_right = client_w - right_margin - self.btnUp.Width
        y = self.lv.Top
        for b in [self.btnUp, self.btnDown, self.btnCopy, self.btnCopyContent, self.btnPDF, self.btnDWG]:
            b.Location = Point(x_right, y); y += b.Height + gap
        y_bottom = client_h - bottom_margin - self.btnNumerate.Height
        self.btnNumerate.Location = Point(client_w - right_margin - self.btnNumerate.Width, y_bottom)
        new_lv_width = (x_right - gap) - self.lv.Left
        if new_lv_width > 260: self.lv.Width = new_lv_width
        new_lv_height = (y_bottom - gap) - self.lv.Top
        if new_lv_height > 180: self.lv.Height = new_lv_height
        # Переставить редактор если открыт
        if self.editBox.Visible and self._edit_ctx is not None:
            self._position_editor(self._edit_ctx['row'], self._edit_ctx['col'])

    def on_resize(self, sender, args):
        self._layout_controls()

    # Построение
    def rebuild_group_values(self):
        sort_param = self.cmbSortParam.SelectedItem if self.cmbSortParam.SelectedItem else DEFAULT_GROUP_PARAM
        values = set()
        for sid, sh in self.sheet_map.items():
            v = get_param_str(sh, sort_param)
            if v: values.add(v)
        values = sorted(list(values), key=natural_key)
        self.cmbGroupVal.Items.Clear()
        for v in values: self.cmbGroupVal.Items.Add(v)
        if self.cmbGroupVal.Items.Count > 0: self.cmbGroupVal.SelectedIndex = 0

    def _make_item(self, index, row, num_param):
        sh = self.sheet_map[row.sheet_id]
        numval = get_param_str(sh, num_param)
        tb_type = get_titleblock_type_name(sh)
        fmtval = get_format_value(sh)
        arr = Array[String]([str(index), row.sheetnum, row.name, numval, tb_type, fmtval])
        it = ListViewItem(arr); it.Tag = row.sheet_id
        return it

    def reload_rows(self, set_initial_order=False):
        sort_param = self.cmbSortParam.SelectedItem if self.cmbSortParam.SelectedItem else DEFAULT_GROUP_PARAM
        group_val  = self.cmbGroupVal.SelectedItem if self.cmbGroupVal.SelectedItem else u""
        num_param  = self.cmbNumParam.SelectedItem if self.cmbNumParam.SelectedItem else DEFAULT_NUM_PARAM
        rows = []
        for r in self.sheets:
            sh = self.sheet_map[r.sheet_id]
            if (get_param_str(sh, sort_param) or u"") == (group_val or u""):
                rows.append(r)
        if set_initial_order or not self.order_ids:
            any_num = any((get_param_str(self.sheet_map[r.sheet_id], num_param) or u"").strip() for r in rows)
            if any_num:
                def key_num(r):
                    nv = (get_param_str(self.sheet_map[r.sheet_id], num_param) or u"").strip()
                    try: return (0, int(nv))
                    except: return (1, natural_key(nv), natural_key(r.sheetnum))
                rows.sort(key=key_num)
            else:
                rows.sort(key=lambda r: natural_key(r.sheetnum))
            self.order_ids = [r.sheet_id for r in rows]
        else:
            id_to_row = {r.sheet_id: r for r in rows}
            rows = [id_to_row[i] for i in self.order_ids if i in id_to_row]
        self.current_rows = rows
        self.col_num.Text = self.cmbNumParam.SelectedItem if self.cmbNumParam.SelectedItem else DEFAULT_NUM_PARAM
        self.lv.BeginUpdate()
        try:
            self.lv.Items.Clear()
            for i, r in enumerate(self.current_rows, 1):
                self.lv.Items.Add(self._make_item(i, r, self.col_num.Text))
            if self.lv.Items.Count > 0: self.lv.Items[0].Selected = True
        finally:
            self.lv.EndUpdate()
        self._hide_editor()

    # События
    def on_sort_param_changed(self, sender, args):
        self.rebuild_group_values(); self.order_ids = []; self.reload_rows(set_initial_order=True)

    def on_group_value_changed(self, sender, args):
        self.order_ids = []; self.reload_rows(set_initial_order=True)

    def on_num_param_changed(self, sender, args):
        self.reload_rows(set_initial_order=False)

    def _selected_index(self):
        return int(self.lv.SelectedIndices[0]) if self.lv.SelectedIndices.Count > 0 else -1

    def on_move_up(self, sender, args):
        idx = self._selected_index()
        if idx <= 0: return
        self.current_rows[idx-1], self.current_rows[idx] = self.current_rows[idx], self.current_rows[idx-1]
        self.order_ids[idx-1], self.order_ids[idx] = self.order_ids[idx], self.order_ids[idx-1]
        self._rerender_keep_selection(idx-1)
        self.lv.Focus()

    def on_move_down(self, sender, args):
        idx = self._selected_index()
        if idx < 0 or idx >= len(self.current_rows)-1: return
        self.current_rows[idx+1], self.current_rows[idx] = self.current_rows[idx], self.current_rows[idx+1]
        self.order_ids[idx+1], self.order_ids[idx] = self.order_ids[idx], self.order_ids[idx+1]
        self._rerender_keep_selection(idx+1)
        self.lv.Focus()

    def on_copy_below(self, sender, args):
        idx = self._selected_index()
        if idx < 0 or idx >= len(self.current_rows): return
        src_row = self.current_rows[idx]
        src_sheet = self.sheet_map[src_row.sheet_id]

        # Тип титульника
        tblks = get_titleblocks_on_sheet(src_sheet)
        symbol_id = None
        if tblks:
            try: symbol_id = tblks[0].Symbol.Id
            except: symbol_id = None
        if symbol_id is None:
            types = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType())
            if types: symbol_id = types[0].Id
        if symbol_id is None:
            MessageBox.Show(u"Не найден тип титульного листа для копирования.", u"Копирование", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return

        t = Transaction(doc, u"Копировать лист"); t.Start(); with_suppress(t)
        try:
            new_sheet = ViewSheet.Create(doc, symbol_id)
            try: new_sheet.Name = (src_sheet.Name or u"") + u" (копия)"
            except: pass
            sort_param = self.cmbSortParam.SelectedItem if self.cmbSortParam.SelectedItem else DEFAULT_GROUP_PARAM
            set_param_str(new_sheet, sort_param, get_param_str(src_sheet, sort_param))
            fmt = get_format_value(src_sheet)
            if fmt != u"": set_format_value(new_sheet, fmt)
        except Exception as ex:
            t.RollBack(); MessageBox.Show(u"Не удалось скопировать лист: " + unicode(ex), u"Копирование", MessageBoxButtons.OK, MessageBoxIcon.Error); return
        t.Commit()

        new_row = SheetRow(new_sheet)
        self.sheets.append(new_row)
        self.sheet_map[new_row.sheet_id] = new_sheet
        insert_at = idx + 1
        self.current_rows.insert(insert_at, new_row)
        self.order_ids.insert(insert_at, new_row.sheet_id)
        self._rerender_keep_selection(insert_at)
        self.lv.Focus()

    def on_copy_with_content(self, sender, args):
        """Копирование листа с содержимым"""
        idx = self._selected_index()
        if idx < 0 or idx >= len(self.current_rows): return
        src_row = self.current_rows[idx]
        src_sheet = self.sheet_map[src_row.sheet_id]
        
        # Собираем содержимое листа
        legends_on_sheet = []  # (viewport, view_name)
        schedules_on_sheet = []  # (schedule_instance, schedule_name)
        viewports_on_sheet = []  # (viewport, view)
        annotations_on_sheet = []  # все аннотации
        text_notes_on_sheet = []
        detail_lines_on_sheet = []
        
        # Получаем все элементы на листе
        collector = FilteredElementCollector(doc, src_sheet.Id).WhereElementIsNotElementType()
        for elem in collector:
            try:
                if isinstance(elem, Viewport):
                    view_id = elem.ViewId
                    view = doc.GetElement(view_id)
                    if view:
                        view_name = view.Name if hasattr(view, 'Name') else u""
                        # Проверяем тип вида
                        try:
                            view_type = view.ViewType
                            if view_type == ViewType.Legend:
                                legends_on_sheet.append((elem, view_name))
                            else:
                                viewports_on_sheet.append((elem, view))
                        except:
                            viewports_on_sheet.append((elem, view))
                elif isinstance(elem, TextNote):
                    text_notes_on_sheet.append(elem)
                elif isinstance(elem, DetailCurve):
                    detail_lines_on_sheet.append(elem)
                elif elem.Category:
                    cat_name = elem.Category.Name if elem.Category else u""
                    # Проверка на спецификации (ScheduleSheetInstance)
                    if "ScheduleSheetInstance" in str(type(elem)):
                        try:
                            sched_id = elem.ScheduleId
                            sched = doc.GetElement(sched_id)
                            if sched:
                                sched_name = sched.Name if hasattr(sched, 'Name') else u""
                                schedules_on_sheet.append((elem, sched_name))
                        except:
                            pass
                    elif elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_TitleBlocks):
                        pass  # Пропускаем титульники
                    elif elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_GenericAnnotation):
                        # Проверяем, не является ли аннотация частью титульника
                        try:
                            if hasattr(elem, 'SuperComponent') and elem.SuperComponent:
                                super_comp = elem.SuperComponent
                                if super_comp.Category and super_comp.Category.Id.IntegerValue == int(BuiltInCategory.OST_TitleBlocks):
                                    continue  # Это вложенный элемент титульника, пропускаем
                        except:
                            pass
                        annotations_on_sheet.append(elem)
            except:
                pass
        
        # Показываем диалог выбора
        dlg = CopyContentDialog(src_sheet, legends_on_sheet, schedules_on_sheet)
        result = dlg.ShowDialog()
        
        if not dlg.result_ok:
            return
        
        # Тип титульника
        tblks = get_titleblocks_on_sheet(src_sheet)
        symbol_id = None
        if tblks:
            try: symbol_id = tblks[0].Symbol.Id
            except: symbol_id = None
        if symbol_id is None:
            types = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType())
            if types: symbol_id = types[0].Id
        if symbol_id is None:
            MessageBox.Show(u"Не найден тип титульного листа.", u"Копирование", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        tg = TransactionGroup(doc, u"Копировать лист с содержимым"); tg.Start()
        
        # Транзакция 1: Создаём лист
        t1 = Transaction(doc, u"Создать лист"); t1.Start(); with_suppress(t1)
        try:
            new_sheet = ViewSheet.Create(doc, symbol_id)
            try: new_sheet.Name = (src_sheet.Name or u"") + u" (копия)"
            except: pass
            sort_param = self.cmbSortParam.SelectedItem if self.cmbSortParam.SelectedItem else DEFAULT_GROUP_PARAM
            set_param_str(new_sheet, sort_param, get_param_str(src_sheet, sort_param))
            fmt = get_format_value(src_sheet)
            if fmt != u"": set_format_value(new_sheet, fmt)
        except Exception as ex:
            t1.RollBack(); tg.RollBack()
            MessageBox.Show(u"Не удалось создать лист: " + unicode(ex), u"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        t1.Commit()
        
        # Транзакция 2: Копируем содержимое
        t2 = Transaction(doc, u"Копировать содержимое"); t2.Start(); with_suppress(t2)
        try:
            from System.Collections.Generic import List
            
            # Копируем текстовые примечания
            if dlg.copy_text_notes and text_notes_on_sheet:
                ids_to_copy = List[ElementId]()
                for tn in text_notes_on_sheet:
                    ids_to_copy.Add(tn.Id)
                if ids_to_copy.Count > 0:
                    ElementTransformUtils.CopyElements(src_sheet, ids_to_copy, new_sheet, None, CopyPasteOptions())
            
            # Копируем линии детализации
            if dlg.copy_detail_lines and detail_lines_on_sheet:
                ids_to_copy = List[ElementId]()
                for dl in detail_lines_on_sheet:
                    ids_to_copy.Add(dl.Id)
                if ids_to_copy.Count > 0:
                    ElementTransformUtils.CopyElements(src_sheet, ids_to_copy, new_sheet, None, CopyPasteOptions())
            
            # Копируем аннотации
            if dlg.copy_annotations and annotations_on_sheet:
                ids_to_copy = List[ElementId]()
                for ann in annotations_on_sheet:
                    ids_to_copy.Add(ann.Id)
                if ids_to_copy.Count > 0:
                    try:
                        ElementTransformUtils.CopyElements(src_sheet, ids_to_copy, new_sheet, None, CopyPasteOptions())
                    except:
                        pass  # Некоторые аннотации могут не копироваться
            
            # Обрабатываем легенды
            legends_errors = []
            for vp, view_name in legends_on_sheet:
                mode = dlg.legends_mode.get(view_name, CopyContentDialog.MODE_NOTHING)
                if mode == CopyContentDialog.MODE_PLACE_SAME:
                    # Размещаем ту же легенду на новом листе
                    try:
                        view_id = vp.ViewId
                        box_center = vp.GetBoxCenter()
                        Viewport.Create(doc, new_sheet.Id, view_id, box_center)
                    except Exception as ex:
                        legends_errors.append(u"Размещение '{0}': {1}".format(view_name, unicode(ex)))
                elif mode == CopyContentDialog.MODE_COPY:
                    # Создаём копию легенды с содержимым и размещаем
                    try:
                        view_id = vp.ViewId
                        legend_view = doc.GetElement(view_id)
                        box_center = vp.GetBoxCenter()
                        
                        # Дублируем легенду с детализацией
                        new_legend_id = None
                        try:
                            new_legend_id = legend_view.Duplicate(ViewDuplicateOption.WithDetailing)
                        except:
                            try:
                                new_legend_id = legend_view.Duplicate(ViewDuplicateOption.Duplicate)
                            except:
                                pass
                        
                        if new_legend_id is not None:
                            new_legend = doc.GetElement(new_legend_id)
                            
                            # Переименовываем
                            try:
                                new_legend.Name = legend_view.Name + u" (копия)"
                            except:
                                pass
                            
                            # Копируем содержимое легенды если оно не скопировалось
                            try:
                                # Получаем все элементы из исходной легенды
                                legend_elements = FilteredElementCollector(doc, view_id).WhereElementIsNotElementType().ToElementIds()
                                if legend_elements.Count > 0:
                                    # Проверяем, есть ли уже элементы в новой легенде
                                    new_elements = FilteredElementCollector(doc, new_legend_id).WhereElementIsNotElementType().ToElementIds()
                                    if new_elements.Count == 0:
                                        # Копируем элементы
                                        ElementTransformUtils.CopyElements(legend_view, legend_elements, new_legend, None, CopyPasteOptions())
                            except:
                                pass
                            
                            # Копируем переопределения графики для элементов
                            try:
                                # Получаем элементы из исходной и новой легенды
                                src_elems = list(FilteredElementCollector(doc, view_id).WhereElementIsNotElementType())
                                new_elems = list(FilteredElementCollector(doc, new_legend_id).WhereElementIsNotElementType())
                                
                                # Создаём словарь соответствия по типу и позиции
                                for src_el in src_elems:
                                    try:
                                        override = legend_view.GetElementOverrides(src_el.Id)
                                        # Ищем соответствующий элемент в новой легенде
                                        for new_el in new_elems:
                                            if new_el.GetTypeId() == src_el.GetTypeId():
                                                new_legend.SetElementOverrides(new_el.Id, override)
                                                new_elems.remove(new_el)
                                                break
                                    except:
                                        pass
                            except:
                                pass
                            
                            # Размещаем на новом листе
                            try:
                                Viewport.Create(doc, new_sheet.Id, new_legend_id, box_center)
                            except Exception as vp_ex:
                                legends_errors.append(u"Размещение копии '{0}': {1}".format(view_name, unicode(vp_ex)))
                        else:
                            # Если дублирование не поддерживается, размещаем ту же легенду
                            Viewport.Create(doc, new_sheet.Id, view_id, box_center)
                            legends_errors.append(u"'{0}': Дублирование не поддерживается, размещена та же легенда".format(view_name))
                    except Exception as ex:
                        legends_errors.append(u"Копирование '{0}': {1}".format(view_name, unicode(ex)))
            
            if legends_errors:
                MessageBox.Show(u"Ошибки при обработке легенд:\n" + u"\n".join(legends_errors), u"Легенды", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            
            # Обрабатываем спецификации
            for sched_inst, sched_name in schedules_on_sheet:
                mode = dlg.schedules_mode.get(sched_name, CopyContentDialog.MODE_NOTHING)
                if mode == CopyContentDialog.MODE_PLACE_SAME:
                    # Размещаем ту же спецификацию
                    try:
                        sched_id = sched_inst.ScheduleId
                        location = sched_inst.Point
                        ScheduleSheetInstance.Create(doc, new_sheet.Id, sched_id, location)
                    except:
                        pass
                elif mode == CopyContentDialog.MODE_COPY:
                    # Создаём копию спецификации и размещаем
                    try:
                        sched_id = sched_inst.ScheduleId
                        sched = doc.GetElement(sched_id)
                        location = sched_inst.Point
                        # Дублируем спецификацию
                        new_sched_id = sched.Duplicate(ViewDuplicateOption.Duplicate)
                        new_sched = doc.GetElement(new_sched_id)
                        try:
                            new_sched.Name = sched.Name + u" (копия)"
                        except:
                            pass
                        ScheduleSheetInstance.Create(doc, new_sheet.Id, new_sched_id, location)
                    except:
                        pass
            
            # Копируем видовые экраны (создаём дубликаты видов)
            if dlg.copy_viewports and viewports_on_sheet:
                for vp, view in viewports_on_sheet:
                    try:
                        box_center = vp.GetBoxCenter()
                        # Дублируем вид
                        new_view_id = view.Duplicate(ViewDuplicateOption.WithDetailing)
                        new_view = doc.GetElement(new_view_id)
                        try:
                            new_view.Name = view.Name + u" (копия)"
                        except:
                            pass
                        # Размещаем на новом листе
                        Viewport.Create(doc, new_sheet.Id, new_view_id, box_center)
                    except:
                        pass
                        
        except Exception as ex:
            t2.RollBack(); tg.RollBack()
            MessageBox.Show(u"Ошибка копирования содержимого: " + unicode(ex), u"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        t2.Commit()
        tg.Assimilate()
        
        # Обновляем данные
        new_row = SheetRow(new_sheet)
        self.sheets.append(new_row)
        self.sheet_map[new_row.sheet_id] = new_sheet
        insert_at = idx + 1
        self.current_rows.insert(insert_at, new_row)
        self.order_ids.insert(insert_at, new_row.sheet_id)
        self._rerender_keep_selection(insert_at)
        self.lv.Focus()

    # In-table edit mechanics
    def on_lv_double_click(self, sender, e):
        hit = self.lv.HitTest(e.Location)
        if hit is None or hit.Item is None or hit.SubItem is None:
            return
        item = hit.Item; sub = hit.SubItem
        row = item.Index
        # find col
        col = -1
        for i in range(item.SubItems.Count):
            if item.SubItems[i] == sub:
                col = i; break
        if col == 4:  # Тип надписи - выпадающий список
            self._start_combo_edit(row, col, sub)
            return
        if col not in (2, 5):  # Имя листа или Формат
            return
        self._start_edit(row, col, sub)

    def _position_editor(self, row, col):
        r = self.lv.Items[row].SubItems[col].Bounds
        self.editBox.Bounds = Rectangle(r.Left + 1, r.Top + 1, max(20, r.Width - 2), max(16, r.Height - 2))
        self.editBox.BringToFront()

    def _start_edit(self, row, col, sub):
        self._edit_ctx = {'row': row, 'col': col}
        self.editBox.Text = sub.Text
        self._position_editor(row, col)
        self.editBox.Visible = True
        self.editBox.Focus()
        self.editBox.SelectAll()

    def _start_combo_edit(self, row, col, sub):
        """Начало редактирования типа надписи через ComboBox"""
        self._hide_editor()  # Скрываем текстовый редактор если открыт
        if not hasattr(self, 'comboBox'):
            self.comboBox = ComboBox()
            self.comboBox.DropDownStyle = ComboBoxStyle.DropDownList
            self.comboBox.Visible = False
            self.comboBox.BackColor = Color.White
            self.comboBox.ForeColor = Color.Black
            self.comboBox.Font = self.lv.Font
            self.comboBox.SelectedIndexChanged += self._combo_selected
            self.comboBox.Leave += self._combo_leave
            self.comboBox.KeyDown += self._combo_keydown
            self.lv.Controls.Add(self.comboBox)
        
        self._combo_ctx = {'row': row, 'col': col}
        
        # Заполняем список типами надписей
        self.comboBox.Items.Clear()
        self._tb_types_dict = get_all_titleblock_types()
        type_names = sorted(self._tb_types_dict.keys(), key=natural_key)
        current_type = sub.Text
        selected_idx = 0
        for i, name in enumerate(type_names):
            self.comboBox.Items.Add(name)
            if name == current_type:
                selected_idx = i
        
        # Позиционируем и показываем
        r = self.lv.Items[row].SubItems[col].Bounds
        self.comboBox.Bounds = Rectangle(r.Left + 1, r.Top + 1, max(20, r.Width - 2), max(16, r.Height - 2))
        self.comboBox.BringToFront()
        if self.comboBox.Items.Count > 0:
            self.comboBox.SelectedIndex = selected_idx
        self.comboBox.Visible = True
        self.comboBox.Focus()
        self.comboBox.DroppedDown = True

    def _combo_selected(self, sender, e):
        """Обработка выбора типа надписи"""
        if not hasattr(self, '_combo_ctx') or self._combo_ctx is None:
            return
        if not self.comboBox.Visible:
            return
        self._commit_combo()

    def _combo_leave(self, sender, e):
        """Скрытие ComboBox при потере фокуса"""
        self._hide_combo()

    def _combo_keydown(self, sender, e):
        if e.KeyCode == Keys.Escape:
            self._hide_combo()
            e.Handled = True

    def _hide_combo(self):
        if hasattr(self, 'comboBox'):
            self.comboBox.Visible = False
        self._combo_ctx = None

    def _commit_combo(self):
        """Сохранение изменения типа надписи"""
        if not hasattr(self, '_combo_ctx') or self._combo_ctx is None:
            return
        if not self.comboBox.Visible:
            return
        
        row = self._combo_ctx['row']
        selected_type = self.comboBox.SelectedItem
        if selected_type is None:
            self._hide_combo()
            return
        
        selected_type = str(selected_type)
        sid = self.current_rows[row].sheet_id
        sh = self.sheet_map[sid]
        current_type = get_titleblock_type_name(sh)
        
        if selected_type != current_type:
            t = Transaction(doc, u"Изменить тип надписи"); t.Start(); with_suppress(t)
            try:
                ok = set_titleblock_type(sh, selected_type, self._tb_types_dict)
                if not ok:
                    MessageBox.Show(u"Не удалось изменить тип надписи.", u"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            except Exception as ex:
                t.RollBack()
                MessageBox.Show(u"Ошибка изменения типа: " + unicode(ex), u"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                self._hide_combo()
                return
            t.Commit()
        
        self._hide_combo()
        self._rerender_keep_selection(row)

    def _hide_editor(self, *args):
        self.editBox.Visible = False
        self._edit_ctx = None

    def _hide_editor_event(self, sender, e):
        self._hide_editor()
        self._hide_combo()

    def _commit_editor(self):
        if not self.editBox.Visible or self._edit_ctx is not None and self._edit_ctx.get('row') is None:
            # defensive
            pass
        if not self.editBox.Visible or self._edit_ctx is None:
            return
        txt = (self.editBox.Text or u"").strip()
        row = self._edit_ctx['row']; col = self._edit_ctx['col']
        sid = self.current_rows[row].sheet_id
        sh = self.sheet_map[sid]

        if col == 2:
            if txt != (sh.Name or u""):
                t = Transaction(doc, u"Переименование листа"); t.Start(); with_suppress(t)
                try:
                    sh.Name = txt
                except Exception as ex:
                    t.RollBack(); MessageBox.Show(u"Не удалось переименовать: " + unicode(ex), u"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); return
                t.Commit()
                self.current_rows[row] = SheetRow(sh)
        elif col == 5:  # Формат (A) - теперь колонка 5
            if txt != u"":
                try:
                    int(txt)
                except:
                    MessageBox.Show(u"Формат (A) должен быть целым числом.", u"Формат", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    return
                t = Transaction(doc, u"Формат A"); t.Start(); with_suppress(t)
                try:
                    ok = set_format_value(sh, txt)
                    if not ok:
                        MessageBox.Show(u"Параметр 'A' недоступен для записи на титульнике.", u"Формат", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                except Exception as ex:
                    t.RollBack(); MessageBox.Show(u"Ошибка установки формата: " + unicode(ex), u"Формат", MessageBoxButtons.OK, MessageBoxIcon.Error); return
                t.Commit()

        self._rerender_keep_selection(row)
        self._hide_editor()

    def _editbox_keydown(self, sender, e):
        if e.KeyCode == Keys.Enter:
            self._commit_editor(); e.Handled = True
        elif e.KeyCode == Keys.Escape:
            self._hide_editor(); e.Handled = True

    def _editbox_leave(self, sender, e):
        self._commit_editor()

    def _rerender_keep_selection(self, new_index):
        num_param = self.cmbNumParam.SelectedItem if self.cmbNumParam.SelectedItem else DEFAULT_NUM_PARAM
        self.lv.BeginUpdate()
        try:
            self.lv.Items.Clear()
            for i, r in enumerate(self.current_rows, 1):
                self.lv.Items.Add(self._make_item(i, r, num_param))
            if 0 <= new_index < self.lv.Items.Count:
                self.lv.Items[new_index].Selected = True
                self.lv.Items[new_index].Focused = True
                self.lv.EnsureVisible(new_index)
        finally:
            self.lv.EndUpdate()

    # Нумерация
    def on_numerate(self, sender, args):
        if not self.current_rows: return
        sort_param = self.cmbSortParam.SelectedItem if self.cmbSortParam.SelectedItem else DEFAULT_GROUP_PARAM
        num_param  = self.cmbNumParam.SelectedItem if self.cmbNumParam.SelectedItem else DEFAULT_NUM_PARAM
        start_at   = int(self.numStart.Value)
        ids_in_order = list(self.order_ids)
        
        # Подсчитываем общее количество листов в альбоме
        total_sheets = len(ids_in_order)
        total_sheets_str = unicode(total_sheets)
        
        plan = []
        for idx, sid in enumerate(ids_in_order, 1):
            sh = self.sheet_map[sid]
            adsk = (get_param_str(sh, sort_param) or u"").strip()
            dvlk = unicode(start_at + idx - 1)
            tmp_num   = adsk + dvlk
            final_num = dvlk + u"_" + adsk
            plan.append((sh, dvlk, adsk, tmp_num, final_num))
        tg = TransactionGroup(doc, u"Нумерация листов"); tg.Start()
        t1 = Transaction(doc, u"Временные номера"); t1.Start(); with_suppress(t1)
        try:
            for (sh, dvlk, adsk, tmp_num, final_num) in plan: set_sheet_number(sh, tmp_num)
        except Exception as ex:
            t1.RollBack(); tg.RollBack(); MessageBox.Show(u"Ошибка на шаге временных номеров: " + unicode(ex), u"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); return
        t1.Commit()
        t2 = Transaction(doc, u"Финальные номера"); t2.Start(); with_suppress(t2)
        try:
            for (sh, dvlk, adsk, tmp_num, final_num) in plan:
                set_sheet_number(sh, final_num)
                set_param_str(sh, num_param, dvlk)
                # Записываем общее количество листов в параметр "Колво листов" в семействе листа
                tblks = get_titleblocks_on_sheet(sh)
                if tblks:
                    for tb in tblks:
                        set_param_on_titleblock(tb, u"Колво листов", total_sheets_str)
        except Exception as ex:
            t2.RollBack(); tg.RollBack(); MessageBox.Show(u"Ошибка на шаге финальных номеров: " + unicode(ex), u"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); return
        t2.Commit()
        tg.Assimilate()
        self.sheets = [SheetRow(s) for s in get_sheets(doc)]
        self.sheet_map = {r.sheet_id: doc.GetElement(ElementId(r.sheet_id)) for r in self.sheets}
        self.order_ids = []
        self.reload_rows(set_initial_order=True)

    def on_pdf_export(self, sender, args):
        """Экспорт листов альбома в PDF через встроенный PDF Export"""
        if not self.current_rows:
            MessageBox.Show(u"Нет листов для экспорта.", u"PDF Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        album_name = self.cmbGroupVal.SelectedItem if self.cmbGroupVal.SelectedItem else u"Листы"
        sheets_count = len(self.current_rows)
        
        # Показываем диалог настроек PDF
        pdf_dialog = PDFExportDialog(album_name, sheets_count)
        if pdf_dialog.ShowDialog() != DialogResult.OK or not pdf_dialog.result_ok:
            return
        
        output_folder = pdf_dialog.selected_folder
        file_name = pdf_dialog.file_name
        combine_pdf = pdf_dialog.combine_pdf
        
        # Собираем листы в порядке отображения
        sheets_to_print = []
        for row in self.current_rows:
            sheet = self.sheet_map.get(row.sheet_id)
            if sheet:
                sheets_to_print.append(sheet)
        
        if not sheets_to_print:
            MessageBox.Show(u"Нет листов для экспорта.", u"PDF Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        try:
            from Autodesk.Revit.DB import PDFExportOptions, ExportPaperFormat
            from System.Collections.Generic import List
            
            # Создаём список ID листов
            sheet_ids = List[ElementId]()
            for sheet in sheets_to_print:
                sheet_ids.Add(sheet.Id)
            
            # Настраиваем опции PDF экспорта
            pdf_options = PDFExportOptions()
            
            # Объединяем в один файл или раздельно
            pdf_options.Combine = combine_pdf
            
            # Автоматический размер бумаги по листу
            pdf_options.PaperFormat = ExportPaperFormat.Default
            
            # Имя файла
            if combine_pdf:
                pdf_options.FileName = file_name
            else:
                # При раздельном экспорте используем имя с номером листа
                pdf_options.FileName = file_name
            
            # Экспортируем
            result = doc.Export(output_folder, sheet_ids, pdf_options)
            
            if result:
                if combine_pdf:
                    full_path = Path.Combine(output_folder, file_name + u".pdf")
                    MessageBox.Show(
                        u"PDF успешно создан!\n\nЛистов: {0}\nФайл: {1}".format(sheets_count, full_path),
                        u"PDF Экспорт",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    )
                else:
                    MessageBox.Show(
                        u"PDF успешно создан!\n\nЛистов: {0}\nПапка: {1}".format(sheets_count, output_folder),
                        u"PDF Экспорт",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    )
            else:
                MessageBox.Show(
                    u"Не удалось экспортировать PDF. Проверьте права доступа к папке.",
                    u"PDF Экспорт",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                
        except Exception as ex:
            MessageBox.Show(u"Ошибка экспорта PDF: " + unicode(ex), u"PDF Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def on_dwg_export(self, sender, args):
        """Экспорт листов альбома в DWG"""
        if not self.current_rows:
            MessageBox.Show(u"Нет листов для экспорта.", u"DWG Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        album_name = self.cmbGroupVal.SelectedItem if self.cmbGroupVal.SelectedItem else u"Листы"
        sheets_count = len(self.current_rows)
        
        # Показываем диалог настроек DWG
        dwg_dialog = DWGExportDialog(album_name, sheets_count)
        if dwg_dialog.ShowDialog() != DialogResult.OK or not dwg_dialog.result_ok:
            return
        
        output_folder = dwg_dialog.selected_folder
        selected_setup = dwg_dialog.selected_setup
        
        # Собираем листы в порядке отображения
        sheets_to_export = []
        for row in self.current_rows:
            sheet = self.sheet_map.get(row.sheet_id)
            if sheet:
                sheets_to_export.append(sheet)
        
        if not sheets_to_export:
            MessageBox.Show(u"Нет листов для экспорта.", u"DWG Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        try:
            from Autodesk.Revit.DB import DWGExportOptions, ExportDWGSettings
            from System.Collections.Generic import List
            
            # Создаём список ID видов (листов)
            view_ids = List[ElementId]()
            for sheet in sheets_to_export:
                view_ids.Add(sheet.Id)
            
            # Получаем настройки экспорта
            dwg_options = DWGExportOptions()
            
            # Ищем пресет по имени
            if selected_setup != u"<По умолчанию>":
                try:
                    collector = FilteredElementCollector(doc).OfClass(ExportDWGSettings)
                    for setup in collector:
                        if setup.Name == selected_setup:
                            dwg_options = setup.GetDWGExportOptions()
                            break
                except:
                    pass
            
            # Отключаем экспорт видов на листах как отдельных файлов
            # 1 лист = 1 DWG файл (MergedViews = True означает что виды на листе включены в файл листа)
            dwg_options.MergedViews = True
            
            # Экспортируем
            result = doc.Export(output_folder, u"", view_ids, dwg_options)
            
            if result:
                MessageBox.Show(
                    u"DWG успешно экспортирован!\n\nЛистов: {0}\nПапка: {1}".format(sheets_count, output_folder),
                    u"DWG Экспорт",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )
            else:
                MessageBox.Show(
                    u"Не удалось экспортировать DWG. Проверьте права доступа к папке.",
                    u"DWG Экспорт",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                
        except Exception as ex:
            MessageBox.Show(u"Ошибка экспорта DWG: " + unicode(ex), u"DWG Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def on_delete_sheet(self, sender, args):
        """Удаление выбранного листа"""
        idx = self._selected_index()
        if idx < 0 or idx >= len(self.current_rows):
            MessageBox.Show(u"Выберите лист для удаления.", u"Удаление", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        src_row = self.current_rows[idx]
        sheet = self.sheet_map[src_row.sheet_id]
        sheet_name = sheet.Name or u""
        sheet_num = get_sheet_number(sheet)
        
        # Подтверждение
        result = MessageBox.Show(
            u"Удалить лист '{0}' ({1})?".format(sheet_name, sheet_num),
            u"Подтверждение удаления",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
        
        if result != DialogResult.Yes:
            return
        
        t = Transaction(doc, u"Удалить лист"); t.Start(); with_suppress(t)
        try:
            doc.Delete(sheet.Id)
        except Exception as ex:
            t.RollBack()
            MessageBox.Show(u"Не удалось удалить лист: " + unicode(ex), u"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        t.Commit()
        
        # Обновляем данные
        del self.sheet_map[src_row.sheet_id]
        self.sheets = [r for r in self.sheets if r.sheet_id != src_row.sheet_id]
        self.current_rows.pop(idx)
        if src_row.sheet_id in self.order_ids:
            self.order_ids.remove(src_row.sheet_id)
        
        # Перерисовываем таблицу
        new_idx = min(idx, len(self.current_rows) - 1)
        if new_idx >= 0:
            self._rerender_keep_selection(new_idx)
        else:
            self.lv.Items.Clear()
        self.lv.Focus()


# --- Запуск ---
form = SheetManagerForm()
form.ShowDialog()
