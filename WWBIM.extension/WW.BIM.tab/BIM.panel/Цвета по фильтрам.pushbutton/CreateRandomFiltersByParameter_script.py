# -*- coding: utf-8 -*-
# Скрипт для pyRevit: создание фильтров вида по текстовому параметру
# и назначение случайного цвета с использованием штриховки Сплошная заливка.
# Расширенный интерфейс с выбором категорий, параметров и предпросмотром цветов.

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Application, Form, Label, ListBox, Panel, Button, CheckedListBox,
    SelectionMode, BorderStyle, FormBorderStyle, FormStartPosition,
    DockStyle, AnchorStyles, MessageBox, MessageBoxButtons, MessageBoxIcon,
    ColorDialog, DialogResult, DrawMode, DrawItemEventArgs
)
from System.Drawing import (
    Size, Point, Color as DrawingColor, Font, FontStyle, 
    SolidBrush, Rectangle, StringFormat, ContentAlignment
)
from System import EventHandler, Array

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FillPatternElement,
    FillPatternTarget,
    OverrideGraphicSettings,
    Color,
    ElementId,
    ParameterFilterElement,
    FilterStringRule,
    FilterStringEquals,
    ParameterValueProvider,
    StorageType,
    FilterRule,
    ElementParameterFilter,
    Transaction,
    View,
    ViewType,
    ViewDuplicateOption,
    Viewport,
    TextNote,
    TextNoteType,
    TextNoteOptions,
    HorizontalTextAlignment,
    FilledRegion,
    FilledRegionType,
    XYZ,
    Line,
    CurveLoop,
    ViewFamilyType,
    ElementTransformUtils,
    CopyPasteOptions,
    Transform,
    Material,
    BuiltInParameter
)
from System.Collections.Generic import List
from System import Random
from pyrevit import revit, forms

doc = revit.doc
view = doc.ActiveView

if view is None:
    forms.alert(u'Нет активного вида.', exitscript=True)


# ============== СБОР ДАННЫХ ==============

def collect_categories_from_view(doc, view):
    """Собирает категории элементов с активного вида"""
    elements_on_view = list(
        FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
    )
    
    categories_dict = {}
    # Исключаем только служебные категории
    excluded_cat_ids = [
        -2000240,  # Уровни
        -2000051,  # Сетки
        -2001320,  # Scope boxes
    ]
    
    for el in elements_on_view:
        cat = el.Category
        if cat is None:
            continue
        
        cid_int = cat.Id.IntegerValue
        if cid_int in excluded_cat_ids:
            continue
        
        try:
            if cat.CategoryType.ToString() == 'Annotation':
                continue
        except:
            pass
        
        cat_name_lower = cat.Name.lower()
        if any(word in cat_name_lower for word in [u'линии', u'lines', u'grid', u'level', u'annotation']):
            continue
        
        try:
            allows_visibility = cat.AllowsVisibilityControl
        except:
            allows_visibility = True
        if not allows_visibility:
            continue
        
        if cid_int not in categories_dict:
            categories_dict[cid_int] = cat
    
    return categories_dict, elements_on_view


def collect_parameters_for_categories(elements_on_view, selected_cat_ids):
    """Собирает текстовые параметры для выбранных категорий"""
    param_names_set = set()
    
    for el in elements_on_view:
        cat = el.Category
        if cat is None or cat.Id not in selected_cat_ids:
            continue
        for p in el.Parameters:
            try:
                if p is None or p.Definition is None:
                    continue
                if p.StorageType != StorageType.String:
                    continue
                name = p.Definition.Name
                if not name:
                    continue
                param_names_set.add(name)
            except:
                continue
    
    return sorted(list(param_names_set))


def collect_parameter_values(doc, elements_on_view, selected_cat_ids, param_name):
    """Собирает уникальные значения параметра"""
    param_id = None
    value_set = set()
    param_cat_ids = set()
    
    for el in elements_on_view:
        cat = el.Category
        if cat is None or cat.Id not in selected_cat_ids:
            continue
        
        # Параметр экземпляра
        param = el.LookupParameter(param_name)
        if param is not None and param.HasValue:
            if param_id is None:
                param_id = param.Id
                if param.StorageType != StorageType.String:
                    continue
            
            val = param.AsString()
            if not val:
                val = param.AsValueString()
            if val:
                val = val.strip()
                if val:
                    value_set.add(val)
                    param_cat_ids.add(cat.Id)
        
        # Параметр типа
        try:
            elem_type = doc.GetElement(el.GetTypeId())
            if elem_type is not None:
                type_param = elem_type.LookupParameter(param_name)
                if type_param is not None and type_param.HasValue:
                    if param_id is None:
                        param_id = type_param.Id
                        if type_param.StorageType != StorageType.String:
                            continue
                    
                    val = type_param.AsString()
                    if not val:
                        val = type_param.AsValueString()
                    if val:
                        val = val.strip()
                        if val:
                            value_set.add(val)
                            param_cat_ids.add(cat.Id)
        except:
            pass
    
    return param_id, sorted(list(value_set)), param_cat_ids


def collect_materials_from_elements(doc, elements_on_view, selected_cat_ids):
    """Собирает уникальные материалы с элементов выбранных категорий"""
    materials_dict = {}  # {material_id: Material}
    
    for el in elements_on_view:
        cat = el.Category
        if cat is None or cat.Id not in selected_cat_ids:
            continue
        
        try:
            # Получаем ID материалов элемента
            material_ids = el.GetMaterialIds(False)  # False = не включать paint
            for mat_id in material_ids:
                if mat_id.IntegerValue > 0 and mat_id not in materials_dict:
                    mat = doc.GetElement(mat_id)
                    if mat is not None:
                        materials_dict[mat_id] = mat
        except:
            pass
        
        # Также проверяем материал из типа
        try:
            elem_type = doc.GetElement(el.GetTypeId())
            if elem_type is not None:
                type_mat_ids = elem_type.GetMaterialIds(False)
                for mat_id in type_mat_ids:
                    if mat_id.IntegerValue > 0 and mat_id not in materials_dict:
                        mat = doc.GetElement(mat_id)
                        if mat is not None:
                            materials_dict[mat_id] = mat
        except:
            pass
    
    # Сортируем по имени
    sorted_materials = sorted(materials_dict.values(), key=lambda m: m.Name)
    return sorted_materials


def get_material_pattern_info(doc, material):
    """Получает информацию о штриховке и цвете материала"""
    pattern_id = None
    color = (128, 128, 128)
    
    try:
        # Пробуем получить Surface Pattern (штриховка поверхности)
        surf_pattern_id = material.SurfacePatternId
        if surf_pattern_id and surf_pattern_id.IntegerValue > 0:
            pattern_id = surf_pattern_id
        
        # Получаем цвет штриховки поверхности
        surf_color = material.SurfacePatternColor
        if surf_color.IsValid:
            color = (surf_color.Red, surf_color.Green, surf_color.Blue)
        else:
            # Fallback на цвет материала
            mat_color = material.Color
            if mat_color.IsValid:
                color = (mat_color.Red, mat_color.Green, mat_color.Blue)
    except:
        try:
            # Для старых версий API
            mat_color = material.Color
            if mat_color.IsValid:
                color = (mat_color.Red, mat_color.Green, mat_color.Blue)
        except:
            pass
    
    return pattern_id, color


def get_solid_fill_pattern(doc):
    """Находит паттерн сплошной заливки"""
    fill_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
    
    for fpe in fill_patterns:
        fp = fpe.GetFillPattern()
        try:
            if fp.IsSolidFill and fp.Target == FillPatternTarget.Drafting:
                return fpe
        except:
            pass
    return None


# ============== ГЕНЕРАЦИЯ ЦВЕТОВ ==============

rand = Random()

# Предустановленные градиенты для циклического переключения
GRADIENT_PRESETS = [
    ((50, 100, 200), (200, 50, 100)),    # Синий -> Красный
    ((50, 180, 80), (180, 50, 180)),     # Зелёный -> Фиолетовый
    ((200, 150, 50), (50, 150, 200)),    # Оранжевый -> Голубой
    ((180, 50, 50), (50, 180, 50)),      # Красный -> Зелёный
    ((100, 50, 180), (180, 180, 50)),    # Фиолетовый -> Жёлтый
    ((50, 180, 180), (180, 80, 50)),     # Бирюзовый -> Коричневый
    ((80, 80, 80), (220, 220, 220)),     # Тёмно-серый -> Светло-серый
    ((200, 100, 150), (100, 200, 150)),  # Розовый -> Мятный
]

gradient_index = [0]  # Используем список для мутабельности в замыкании

def generate_random_colors(count):
    """Генерирует список случайных цветов"""
    colors = []
    for _ in range(count):
        r = rand.Next(50, 230)
        g = rand.Next(50, 230)
        b = rand.Next(50, 230)
        colors.append((r, g, b))
    return colors


def generate_gradient_colors(count, start_color=None, end_color=None):
    """Генерирует градиентные цвета. Если цвета не указаны, берёт следующий пресет."""
    if start_color is None or end_color is None:
        preset = GRADIENT_PRESETS[gradient_index[0] % len(GRADIENT_PRESETS)]
        gradient_index[0] += 1
        start_color, end_color = preset
    
    if count <= 1:
        return [start_color]
    
    colors = []
    for i in range(count):
        t = float(i) / (count - 1)
        r = int(start_color[0] + t * (end_color[0] - start_color[0]))
        g = int(start_color[1] + t * (end_color[1] - start_color[1]))
        b = int(start_color[2] + t * (end_color[2] - start_color[2]))
        colors.append((r, g, b))
    return colors


# ============== WINFORMS ИНТЕРФЕЙС ==============

class ColorListBox(ListBox):
    """Кастомный ListBox с отображением цветов"""
    
    def __init__(self):
        super(ColorListBox, self).__init__()
        self.DrawMode = DrawMode.OwnerDrawFixed
        self.ItemHeight = 24
        self.colors_dict = {}  # {value_name: (r, g, b)}
        self.DrawItem += self.on_draw_item
    
    def on_draw_item(self, sender, e):
        if e.Index < 0:
            return
        
        e.DrawBackground()
        
        item_text = str(self.Items[e.Index])
        color_tuple = self.colors_dict.get(item_text, (200, 200, 200))
        
        # Рисуем цветной прямоугольник
        color_rect = Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, 40, e.Bounds.Height - 4)
        brush = SolidBrush(DrawingColor.FromArgb(color_tuple[0], color_tuple[1], color_tuple[2]))
        e.Graphics.FillRectangle(brush, color_rect)
        
        # Рисуем текст
        text_brush = SolidBrush(DrawingColor.Black)
        text_point = Point(e.Bounds.X + 48, e.Bounds.Y + 3)
        e.Graphics.DrawString(item_text, e.Font, text_brush, float(text_point.X), float(text_point.Y))
        
        e.DrawFocusRectangle()
        
        brush.Dispose()
        text_brush.Dispose()
    
    def set_colors(self, colors_dict):
        self.colors_dict = colors_dict
        self.Refresh()


class FilterColorForm(Form):
    """Главное окно плагина"""
    
    def __init__(self, doc, view, categories_dict, elements_on_view):
        self.doc = doc
        self.view = view
        self.categories_dict = categories_dict
        self.elements_on_view = elements_on_view
        self.sorted_cats = sorted(categories_dict.values(), key=lambda c: c.Name)
        
        self.selected_cat_ids = []
        self.param_names = []
        self.selected_param = None
        self.values_list = []
        self.colors_dict = {}  # {value: (r, g, b)}
        self.param_id = None
        self.param_cat_ids = set()
        
        self.result_action = None  # 'filters', 'temp_view', 'legend'
        
        self.init_ui()
    
    def init_ui(self):
        self.Text = u'Цвета по фильтрам'
        self.Size = Size(950, 620)
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.StartPosition = FormStartPosition.CenterScreen
        self.MinimumSize = Size(900, 550)
        
        # === Левая панель: Категории ===
        self.lbl_categories = Label()
        self.lbl_categories.Text = u'Категории:'
        self.lbl_categories.Location = Point(10, 10)
        self.lbl_categories.Size = Size(260, 20)
        self.Controls.Add(self.lbl_categories)
        
        self.lst_categories = CheckedListBox()
        self.lst_categories.Location = Point(10, 35)
        self.lst_categories.Size = Size(260, 450)
        self.lst_categories.CheckOnClick = True
        self.lst_categories.ScrollAlwaysVisible = True
        self.lst_categories.IntegralHeight = False
        self.lst_categories.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
        self.lst_categories.ItemCheck += self.on_category_check
        self.Controls.Add(self.lst_categories)
        
        for cat in self.sorted_cats:
            self.lst_categories.Items.Add(cat.Name)
        
        # === Средняя панель: Параметры ===
        self.lbl_params = Label()
        self.lbl_params.Text = u'Параметры:'
        self.lbl_params.Location = Point(280, 10)
        self.lbl_params.Size = Size(260, 20)
        self.Controls.Add(self.lbl_params)
        
        self.lst_params = ListBox()
        self.lst_params.Location = Point(280, 35)
        self.lst_params.Size = Size(260, 450)
        self.lst_params.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
        self.lst_params.SelectedIndexChanged += self.on_param_selected
        self.Controls.Add(self.lst_params)
        
        # === Правая панель: Значения с цветами ===
        self.lbl_values = Label()
        self.lbl_values.Text = u'Значения и цвета:'
        self.lbl_values.Location = Point(550, 10)
        self.lbl_values.Size = Size(380, 20)
        self.Controls.Add(self.lbl_values)
        
        self.lst_values = ColorListBox()
        self.lst_values.Location = Point(550, 35)
        self.lst_values.Size = Size(380, 350)
        self.lst_values.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right
        self.lst_values.SelectionMode = SelectionMode.One
        self.lst_values.DoubleClick += self.on_value_double_click
        self.Controls.Add(self.lst_values)
        
        # === Кнопки управления цветами (под блоком цветов) ===
        color_btn_y = 395
        btn_height = 30
        btn_width = 120
        
        self.btn_gradient = Button()
        self.btn_gradient.Text = u'Градиент'
        self.btn_gradient.Location = Point(550, color_btn_y)
        self.btn_gradient.Size = Size(btn_width, btn_height)
        self.btn_gradient.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        self.btn_gradient.Click += self.on_gradient_click
        self.Controls.Add(self.btn_gradient)
        
        self.btn_random = Button()
        self.btn_random.Text = u'Случайные'
        self.btn_random.Location = Point(680, color_btn_y)
        self.btn_random.Size = Size(btn_width, btn_height)
        self.btn_random.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        self.btn_random.Click += self.on_random_click
        self.Controls.Add(self.btn_random)
        
        self.btn_legend = Button()
        self.btn_legend.Text = u'Легенда (цвета)'
        self.btn_legend.Location = Point(810, color_btn_y)
        self.btn_legend.Size = Size(btn_width, btn_height)
        self.btn_legend.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_legend.Click += self.on_legend_click
        self.Controls.Add(self.btn_legend)
        
        # Второй ряд кнопок цветов
        color_btn_y2 = 430
        
        self.btn_material_legend = Button()
        self.btn_material_legend.Text = u'Легенда (материалы)'
        self.btn_material_legend.Location = Point(550, color_btn_y2)
        self.btn_material_legend.Size = Size(150, btn_height)
        self.btn_material_legend.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        self.btn_material_legend.Click += self.on_material_legend_click
        self.Controls.Add(self.btn_material_legend)
        
        # === Нижняя панель: Основные действия ===
        bottom_y = 500
        
        self.btn_filters = Button()
        self.btn_filters.Text = u'Создать фильтры'
        self.btn_filters.Location = Point(10, bottom_y)
        self.btn_filters.Size = Size(140, 35)
        self.btn_filters.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        self.btn_filters.Click += self.on_filters_click
        self.Controls.Add(self.btn_filters)
        
        self.btn_temp_view = Button()
        self.btn_temp_view.Text = u'Переопределить графику'
        self.btn_temp_view.Location = Point(160, bottom_y)
        self.btn_temp_view.Size = Size(170, 35)
        self.btn_temp_view.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        self.btn_temp_view.Click += self.on_temp_view_click
        self.Controls.Add(self.btn_temp_view)
        
        self.btn_reset = Button()
        self.btn_reset.Text = u'Сбросить графику'
        self.btn_reset.Location = Point(340, bottom_y)
        self.btn_reset.Size = Size(140, 35)
        self.btn_reset.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        self.btn_reset.Click += self.on_reset_click
        self.Controls.Add(self.btn_reset)
        
        # Статусная строка
        self.lbl_status = Label()
        self.lbl_status.Text = u'Выберите категории для начала работы'
        self.lbl_status.Location = Point(10, 545)
        self.lbl_status.Size = Size(920, 25)
        self.lbl_status.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.lbl_status)
    
    def on_category_check(self, sender, e):
        """Обработчик выбора категорий"""
        # Получаем выбранные категории после изменения
        def update_after_check():
            self.selected_cat_ids = []
            for i in range(self.lst_categories.Items.Count):
                if self.lst_categories.GetItemChecked(i):
                    cat_name = str(self.lst_categories.Items[i])
                    for cat in self.sorted_cats:
                        if cat.Name == cat_name:
                            self.selected_cat_ids.append(cat.Id)
                            break
            
            # Обновляем список параметров
            self.update_parameters_list()
        
        # Используем BeginInvoke для отложенного вызова
        from System import Action
        self.BeginInvoke(Action(update_after_check))
    
    def update_parameters_list(self):
        """Обновляет список параметров для выбранных категорий"""
        self.lst_params.Items.Clear()
        self.lst_values.Items.Clear()
        self.values_list = []
        self.colors_dict = {}
        
        if not self.selected_cat_ids:
            self.lbl_status.Text = u'Выберите категории'
            return
        
        self.param_names = collect_parameters_for_categories(
            self.elements_on_view, self.selected_cat_ids
        )
        
        for name in self.param_names:
            self.lst_params.Items.Add(name)
        
        self.lbl_status.Text = u'Найдено {0} параметров. Выберите параметр.'.format(len(self.param_names))
    
    def on_param_selected(self, sender, e):
        """Обработчик выбора параметра"""
        if self.lst_params.SelectedIndex < 0:
            return
        
        self.selected_param = str(self.lst_params.SelectedItem)
        self.update_values_list()
    
    def update_values_list(self):
        """Обновляет список значений параметра"""
        self.lst_values.Items.Clear()
        self.values_list = []
        self.colors_dict = {}
        
        if not self.selected_param:
            return
        
        self.param_id, self.values_list, self.param_cat_ids = collect_parameter_values(
            self.doc, self.elements_on_view, self.selected_cat_ids, self.selected_param
        )
        
        if not self.values_list:
            self.lbl_status.Text = u'Нет значений для параметра "{0}"'.format(self.selected_param)
            return
        
        # Генерируем случайные цвета по умолчанию
        colors = generate_random_colors(len(self.values_list))
        for i, val in enumerate(self.values_list):
            self.colors_dict[val] = colors[i]
            self.lst_values.Items.Add(val)
        
        self.lst_values.set_colors(self.colors_dict)
        self.lbl_status.Text = u'Найдено {0} уникальных значений. Двойной клик для изменения цвета.'.format(
            len(self.values_list)
        )
    
    def on_gradient_click(self, sender, e):
        """Применить градиентные цвета (при каждом нажатии - новый градиент)"""
        if not self.values_list:
            MessageBox.Show(u'Сначала выберите параметр', u'Внимание', 
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        colors = generate_gradient_colors(len(self.values_list))
        for i, val in enumerate(self.values_list):
            self.colors_dict[val] = colors[i]
        
        self.lst_values.set_colors(self.colors_dict)
        preset_num = (gradient_index[0] - 1) % len(GRADIENT_PRESETS) + 1
        self.lbl_status.Text = u'Применён градиент {0} из {1}. Нажмите ещё раз для следующего.'.format(
            preset_num, len(GRADIENT_PRESETS)
        )
    
    def on_random_click(self, sender, e):
        """Применить случайные цвета"""
        if not self.values_list:
            MessageBox.Show(u'Сначала выберите параметр', u'Внимание',
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        colors = generate_random_colors(len(self.values_list))
        for i, val in enumerate(self.values_list):
            self.colors_dict[val] = colors[i]
        
        self.lst_values.set_colors(self.colors_dict)
        self.lbl_status.Text = u'Применены случайные цвета'
    
    def on_value_double_click(self, sender, e):
        """Изменить цвет для выбранного значения"""
        if self.lst_values.SelectedIndex < 0:
            return
        
        val = str(self.lst_values.SelectedItem)
        current_color = self.colors_dict.get(val, (200, 200, 200))
        
        color_dialog = ColorDialog()
        color_dialog.Color = DrawingColor.FromArgb(current_color[0], current_color[1], current_color[2])
        
        if color_dialog.ShowDialog() == DialogResult.OK:
            new_color = color_dialog.Color
            self.colors_dict[val] = (new_color.R, new_color.G, new_color.B)
            self.lst_values.set_colors(self.colors_dict)
    
    def on_legend_click(self, sender, e):
        """Создать легенду"""
        if not self.validate_selection():
            return
        self.result_action = 'legend'
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_material_legend_click(self, sender, e):
        """Создать легенду по материалам"""
        if not self.selected_cat_ids:
            MessageBox.Show(u'Выберите хотя бы одну категорию', u'Внимание',
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.result_action = 'material_legend'
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_filters_click(self, sender, e):
        """Создать фильтры"""
        if not self.validate_selection():
            return
        self.result_action = 'filters'
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_temp_view_click(self, sender, e):
        """Временный вид"""
        if not self.validate_selection():
            return
        self.result_action = 'temp_view'
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def on_reset_click(self, sender, e):
        """Сбросить переопределения графики"""
        if not self.selected_cat_ids:
            MessageBox.Show(u'Выберите хотя бы одну категорию', u'Внимание',
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.result_action = 'reset_graphics'
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def validate_selection(self):
        """Проверка выбора"""
        if not self.selected_cat_ids:
            MessageBox.Show(u'Выберите хотя бы одну категорию', u'Внимание',
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return False
        
        if not self.selected_param:
            MessageBox.Show(u'Выберите параметр', u'Внимание',
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return False
        
        if not self.values_list:
            MessageBox.Show(u'Нет значений для выбранного параметра', u'Внимание',
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return False
        
        if not self.param_id:
            MessageBox.Show(u'Не удалось определить ID параметра', u'Внимание',
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return False
        
        return True


# ============== СОЗДАНИЕ ФИЛЬТРОВ ==============

def create_filters(doc, view, param_name, param_id, values_list, colors_dict, param_cat_ids, solid_fill):
    """Создаёт фильтры вида с заданными цветами"""
    
    catid_list = List[ElementId]()
    for cid in param_cat_ids:
        catid_list.Add(cid)
    
    existing_filters = {}
    for pf in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        existing_filters[pf.Name] = pf
    
    username = doc.Application.Username
    pvp = ParameterValueProvider(param_id)
    
    created_count = 0
    updated_count = 0
    applied_count = 0
    
    with revit.Transaction(u'Создание фильтров по параметру "{0}"'.format(param_name)):
        for val in values_list:
            safe_val = val.replace('.', '_').replace('/', '_').replace('\\', '_')
            filter_name = u'{0}_{1}_{2}'.format(param_name, safe_val, username)
            color_tuple = colors_dict.get(val, (128, 128, 128))
            color = Color(color_tuple[0], color_tuple[1], color_tuple[2])
            
            try:
                rules_list = List[FilterRule]()
                rule = FilterStringRule(pvp, FilterStringEquals(), val)
                rules_list.Add(rule)
                elem_filter = ElementParameterFilter(rules_list)
            except:
                continue
            
            pf = existing_filters.get(filter_name)
            if pf:
                try:
                    pf.SetCategories(catid_list)
                    pf.SetElementFilter(elem_filter)
                    updated_count += 1
                except:
                    continue
            else:
                try:
                    pf = ParameterFilterElement.Create(doc, filter_name, catid_list, elem_filter)
                    existing_filters[filter_name] = pf
                    created_count += 1
                except:
                    continue
            
            try:
                if not view.IsFilterApplied(pf.Id):
                    view.AddFilter(pf.Id)
            except:
                try:
                    view.AddFilter(pf.Id)
                except:
                    pass
            
            ogs = OverrideGraphicSettings()
            try:
                ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                ogs.SetSurfaceForegroundPatternColor(color)
                ogs.SetSurfaceForegroundPatternVisible(True)
                ogs.SetCutForegroundPatternId(solid_fill.Id)
                ogs.SetCutForegroundPatternColor(color)
                ogs.SetCutForegroundPatternVisible(True)
            except:
                try:
                    ogs.SetProjectionFillPatternId(solid_fill.Id)
                    ogs.SetProjectionFillColor(color)
                    ogs.SetCutFillPatternId(solid_fill.Id)
                    ogs.SetCutFillColor(color)
                except:
                    pass
            
            try:
                view.SetFilterOverrides(pf.Id, ogs)
                applied_count += 1
            except:
                pass
    
    return created_count, updated_count, applied_count


def apply_temp_view_overrides(doc, view, param_name, elements_on_view, selected_cat_ids, colors_dict, solid_fill):
    """Применяет временные переопределения графики к элементам"""
    
    applied_count = 0
    
    with revit.Transaction(u'Временные цвета по "{0}"'.format(param_name)):
        for el in elements_on_view:
            cat = el.Category
            if cat is None or cat.Id not in selected_cat_ids:
                continue
            
            # Ищем значение параметра
            val = None
            param = el.LookupParameter(param_name)
            if param is not None and param.HasValue:
                val = param.AsString()
                if not val:
                    val = param.AsValueString()
            
            if not val:
                try:
                    elem_type = doc.GetElement(el.GetTypeId())
                    if elem_type:
                        type_param = elem_type.LookupParameter(param_name)
                        if type_param and type_param.HasValue:
                            val = type_param.AsString()
                            if not val:
                                val = type_param.AsValueString()
                except:
                    pass
            
            if not val or val.strip() not in colors_dict:
                continue
            
            val = val.strip()
            color_tuple = colors_dict[val]
            color = Color(color_tuple[0], color_tuple[1], color_tuple[2])
            
            ogs = OverrideGraphicSettings()
            try:
                ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                ogs.SetSurfaceForegroundPatternColor(color)
                ogs.SetSurfaceForegroundPatternVisible(True)
                ogs.SetCutForegroundPatternId(solid_fill.Id)
                ogs.SetCutForegroundPatternColor(color)
                ogs.SetCutForegroundPatternVisible(True)
            except:
                try:
                    ogs.SetProjectionFillPatternId(solid_fill.Id)
                    ogs.SetProjectionFillColor(color)
                    ogs.SetCutFillPatternId(solid_fill.Id)
                    ogs.SetCutFillColor(color)
                except:
                    pass
            
            try:
                view.SetElementOverrides(el.Id, ogs)
                applied_count += 1
            except:
                pass
    
    return applied_count


def reset_element_overrides(doc, view, elements_on_view, selected_cat_ids):
    """Сбрасывает все переопределения графики для элементов выбранных категорий"""
    reset_count = 0
    
    with revit.Transaction(u'Сброс переопределений графики'):
        # Пустые настройки = сброс
        ogs = OverrideGraphicSettings()
        
        for el in elements_on_view:
            cat = el.Category
            if cat is None or cat.Id not in selected_cat_ids:
                continue
            
            try:
                view.SetElementOverrides(el.Id, ogs)
                reset_count += 1
            except:
                pass
    
    return reset_count


def create_legend_view(doc, category_names, param_name, values_list, colors_dict, solid_fill):
    """Создаёт вид-легенду с цветовыми прямоугольниками и подписями"""
    
    # Ищем существующий вид-легенду для дублирования
    existing_legend = None
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if v.ViewType == ViewType.Legend:
                existing_legend = v
                break
        except:
            continue
    
    if existing_legend is None:
        return None, u'В проекте нет ни одной легенды для дублирования.\nСоздайте легенду вручную.'
    
    # Ищем тип текста
    text_note_type = None
    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):
        text_note_type = tnt
        break
    
    if text_note_type is None:
        return None, u'Не найден тип текстовой заметки'
    
    # Ищем тип FilledRegion
    filled_region_type = None
    for frt in FilteredElementCollector(doc).OfClass(FilledRegionType):
        filled_region_type = frt
        break
    
    if filled_region_type is None:
        return None, u'Не найден тип заполненной области'
    
    # Формируем имя легенды
    cat_names_short = ', '.join(category_names[:2])
    if len(category_names) > 2:
        cat_names_short += u'...'
    legend_name = u'{0} / {1}'.format(cat_names_short, param_name)
    
    # Проверяем существующие виды с таким именем
    existing_views = {}
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            existing_views[v.Name] = v
        except:
            pass
    
    # Если легенда существует - добавляем суффикс
    base_name = legend_name
    counter = 1
    while legend_name in existing_views:
        legend_name = u'{0} ({1})'.format(base_name, counter)
        counter += 1
    
    legend_view = None
    
    with revit.TransactionGroup(u'Создание легенды "{0}"'.format(legend_name)):
        # Транзакция 1: Дублируем легенду
        with revit.Transaction(u'Дублирование легенды'):
            try:
                new_view_id = existing_legend.Duplicate(ViewDuplicateOption.Duplicate)
                legend_view = doc.GetElement(new_view_id)
                # Переименовываем
                i = 1
                while True:
                    try:
                        legend_view.Name = legend_name if i == 1 else u'{0} - {1}'.format(legend_name, i)
                        break
                    except:
                        i += 1
                        if i > 100:
                            break
            except Exception as ex:
                return None, u'Ошибка дублирования легенды: {0}'.format(str(ex))
        
        # Транзакция 2: Создаём графику легенды
        with revit.Transaction(u'Создание графики легенды'):
            # Ищем FilledRegionType со сплошной заливкой
            filled_type = None
            all_filled_types = FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements()
            for frt in all_filled_types:
                try:
                    pattern = doc.GetElement(frt.ForegroundPatternId)
                    if pattern is not None:
                        if pattern.GetFillPattern().IsSolidFill:
                            if frt.ForegroundPatternColor.IsValid:
                                filled_type = frt
                                break
                except:
                    continue
            
            # Если не нашли подходящий тип, берём первый
            if filled_type is None and len(all_filled_types) > 0:
                filled_type = all_filled_types[0]
            
            if filled_type is None:
                return None, u'Не найден тип заполненной области'
            
            # Создаём заголовок
            title_loc = XYZ(0, 0, 0)
            title_text = u'{0} / {1}'.format(cat_names_short, param_name)
            
            try:
                legend_title = TextNote.Create(doc, legend_view.Id, title_loc, title_text, text_note_type.Id)
            except:
                return None, u'Ошибка создания заголовка'
            
            doc.Regenerate()
            
            # Получаем размеры заголовка
            title_bbox = legend_title.get_BoundingBox(legend_view)
            offset = (title_bbox.Max.Y - title_bbox.Min.Y) * 1.5
            
            list_max_X = []
            list_y = []
            height = 0
            fin_coord_y = title_bbox.Min.Y - offset
            
            # Создаём текстовые подписи для каждого значения
            for idx, val in enumerate(values_list):
                if idx == 0:
                    point = XYZ(0, fin_coord_y, 0)
                else:
                    point = XYZ(0, fin_coord_y, 0)
                
                try:
                    new_text = TextNote.Create(doc, legend_view.Id, point, val, text_note_type.Id)
                except:
                    continue
                
                doc.Regenerate()
                
                prev_bbox = new_text.get_BoundingBox(legend_view)
                text_offset = (prev_bbox.Max.Y - prev_bbox.Min.Y) * 0.25
                fin_coord_y = prev_bbox.Min.Y - text_offset
                list_max_X.append(prev_bbox.Max.X)
                list_y.append(prev_bbox.Min.Y)
                height = (prev_bbox.Max.Y - prev_bbox.Min.Y) * 0.8
            
            if not list_max_X:
                return None, u'Не удалось создать текстовые элементы'
            
            # Начальная X координата для прямоугольников (справа от текста)
            ini_x = max(list_max_X) + height * 0.5
            
            # Настройки переопределения графики
            ogs = OverrideGraphicSettings()
            
            # Создаём цветные прямоугольники
            for idx, coord_y in enumerate(list_y):
                color_tuple = colors_dict.get(values_list[idx], (128, 128, 128))
                
                # Координаты прямоугольника
                point0 = XYZ(ini_x, coord_y, 0)
                point1 = XYZ(ini_x, coord_y + height, 0)
                point2 = XYZ(ini_x + height * 2, coord_y + height, 0)
                point3 = XYZ(ini_x + height * 2, coord_y, 0)
                
                line01 = Line.CreateBound(point0, point1)
                line12 = Line.CreateBound(point1, point2)
                line23 = Line.CreateBound(point2, point3)
                line30 = Line.CreateBound(point3, point0)
                
                curveLoops = CurveLoop()
                curveLoops.Append(line01)
                curveLoops.Append(line12)
                curveLoops.Append(line23)
                curveLoops.Append(line30)
                
                list_curveLoops = List[CurveLoop]()
                list_curveLoops.Add(curveLoops)
                
                try:
                    reg = FilledRegion.Create(doc, filled_type.Id, legend_view.Id, list_curveLoops)
                    
                    # Назначаем цвет через переопределение графики
                    color = Color(color_tuple[0], color_tuple[1], color_tuple[2])
                    ogs.SetSurfaceForegroundPatternColor(color)
                    ogs.SetCutForegroundPatternColor(color)
                    if solid_fill is not None:
                        ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                        ogs.SetCutForegroundPatternId(solid_fill.Id)
                    
                    legend_view.SetElementOverrides(reg.Id, ogs)
                except:
                    pass
    
    return legend_view, None


def create_material_legend_view(doc, category_names, materials_list, solid_fill):
    """Создаёт вид-легенду с штриховками материалов"""
    
    # Ищем существующий вид-легенду для дублирования
    existing_legend = None
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if v.ViewType == ViewType.Legend:
                existing_legend = v
                break
        except:
            continue
    
    if existing_legend is None:
        return None, u'В проекте нет ни одной легенды для дублирования.\nСоздайте легенду вручную.'
    
    # Ищем тип текста
    text_note_type = None
    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):
        text_note_type = tnt
        break
    
    if text_note_type is None:
        return None, u'Не найден тип текстовой заметки'
    
    # Формируем имя легенды
    cat_names_short = ', '.join(category_names[:2])
    if len(category_names) > 2:
        cat_names_short += u'...'
    legend_name = u'{0} / Материалы'.format(cat_names_short)
    
    # Проверяем существующие виды с таким именем
    existing_views = {}
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            existing_views[v.Name] = v
        except:
            pass
    
    legend_view = None
    
    with revit.TransactionGroup(u'Создание легенды материалов'):
        # Транзакция 1: Дублируем легенду
        with revit.Transaction(u'Дублирование легенды'):
            try:
                new_view_id = existing_legend.Duplicate(ViewDuplicateOption.Duplicate)
                legend_view = doc.GetElement(new_view_id)
                # Переименовываем
                i = 1
                while True:
                    try:
                        legend_view.Name = legend_name if i == 1 else u'{0} - {1}'.format(legend_name, i)
                        break
                    except:
                        i += 1
                        if i > 100:
                            break
            except Exception as ex:
                return None, u'Ошибка дублирования легенды: {0}'.format(str(ex))
        
        # Транзакция 2: Создаём графику легенды
        with revit.Transaction(u'Создание графики легенды материалов'):
            # Создаём заголовок
            title_loc = XYZ(0, 0, 0)
            title_text = u'{0} / Материалы'.format(cat_names_short)
            
            try:
                legend_title = TextNote.Create(doc, legend_view.Id, title_loc, title_text, text_note_type.Id)
            except:
                return None, u'Ошибка создания заголовка'
            
            doc.Regenerate()
            
            # Получаем размеры заголовка
            title_bbox = legend_title.get_BoundingBox(legend_view)
            offset = (title_bbox.Max.Y - title_bbox.Min.Y) * 1.5
            
            list_max_X = []
            list_y = []
            materials_data = []  # [(material, pattern_id, color), ...]
            height = 0
            fin_coord_y = title_bbox.Min.Y - offset
            
            # Создаём текстовые подписи для каждого материала
            for idx, mat in enumerate(materials_list):
                mat_name = mat.Name
                pattern_id, color = get_material_pattern_info(doc, mat)
                materials_data.append((mat, pattern_id, color))
                
                point = XYZ(0, fin_coord_y, 0)
                
                try:
                    new_text = TextNote.Create(doc, legend_view.Id, point, mat_name, text_note_type.Id)
                except:
                    continue
                
                doc.Regenerate()
                
                prev_bbox = new_text.get_BoundingBox(legend_view)
                text_offset = (prev_bbox.Max.Y - prev_bbox.Min.Y) * 0.25
                fin_coord_y = prev_bbox.Min.Y - text_offset
                list_max_X.append(prev_bbox.Max.X)
                list_y.append(prev_bbox.Min.Y)
                height = (prev_bbox.Max.Y - prev_bbox.Min.Y) * 0.8
            
            if not list_max_X:
                return None, u'Не удалось создать текстовые элементы'
            
            # Начальная X координата для прямоугольников (справа от текста)
            ini_x = max(list_max_X) + height * 0.5
            
            # Ищем FilledRegionType для использования
            filled_type = None
            all_filled_types = FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements()
            if len(all_filled_types) > 0:
                filled_type = all_filled_types[0]
            
            if filled_type is None:
                return None, u'Не найден тип заполненной области'
            
            # Настройки переопределения графики
            ogs = OverrideGraphicSettings()
            
            # Создаём прямоугольники с штриховками материалов
            for idx, coord_y in enumerate(list_y):
                mat, pattern_id, color_tuple = materials_data[idx]
                
                # Координаты прямоугольника
                point0 = XYZ(ini_x, coord_y, 0)
                point1 = XYZ(ini_x, coord_y + height, 0)
                point2 = XYZ(ini_x + height * 2, coord_y + height, 0)
                point3 = XYZ(ini_x + height * 2, coord_y, 0)
                
                line01 = Line.CreateBound(point0, point1)
                line12 = Line.CreateBound(point1, point2)
                line23 = Line.CreateBound(point2, point3)
                line30 = Line.CreateBound(point3, point0)
                
                curveLoops = CurveLoop()
                curveLoops.Append(line01)
                curveLoops.Append(line12)
                curveLoops.Append(line23)
                curveLoops.Append(line30)
                
                list_curveLoops = List[CurveLoop]()
                list_curveLoops.Add(curveLoops)
                
                try:
                    reg = FilledRegion.Create(doc, filled_type.Id, legend_view.Id, list_curveLoops)
                    
                    # Назначаем штриховку и цвет материала
                    color = Color(color_tuple[0], color_tuple[1], color_tuple[2])
                    ogs.SetSurfaceForegroundPatternColor(color)
                    ogs.SetCutForegroundPatternColor(color)
                    
                    # Используем паттерн материала если есть, иначе сплошную заливку
                    if pattern_id is not None and pattern_id.IntegerValue > 0:
                        ogs.SetSurfaceForegroundPatternId(pattern_id)
                        ogs.SetCutForegroundPatternId(pattern_id)
                    elif solid_fill is not None:
                        ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                        ogs.SetCutForegroundPatternId(solid_fill.Id)
                    
                    legend_view.SetElementOverrides(reg.Id, ogs)
                except:
                    pass
    
    return legend_view, None


# ============== ГЛАВНАЯ ЛОГИКА ==============

# Сбор данных
categories_dict, elements_on_view = collect_categories_from_view(doc, view)

if not categories_dict:
    forms.alert(u'На активном виде нет элементов для создания фильтров.', exitscript=True)

solid_fill = get_solid_fill_pattern(doc)
if solid_fill is None:
    forms.alert(u'Не найден штриховочный паттерн "Сплошная заливка".', exitscript=True)

# Показываем форму
form = FilterColorForm(doc, view, categories_dict, elements_on_view)
result = form.ShowDialog()

if result != DialogResult.OK:
    # Пользователь отменил
    import sys
    sys.exit()

# Получаем данные из формы
param_name = form.selected_param
param_id = form.param_id
values_list = form.values_list
colors_dict = form.colors_dict
param_cat_ids = form.param_cat_ids
action = form.result_action
selected_cat_ids = form.selected_cat_ids

# Выполняем выбранное действие
if action == 'filters':
    created, updated, applied = create_filters(
        doc, view, param_name, param_id, values_list, colors_dict, param_cat_ids, solid_fill
    )
    # Успех - не показываем сообщение

elif action == 'temp_view':
    applied = apply_temp_view_overrides(
        doc, view, param_name, elements_on_view, selected_cat_ids, colors_dict, solid_fill
    )
    # Успех - не показываем сообщение

elif action == 'reset_graphics':
    reset_count = reset_element_overrides(doc, view, elements_on_view, selected_cat_ids)
    # Успех - не показываем сообщение

elif action == 'legend':
    # Получаем имена выбранных категорий
    category_names = []
    for cat in form.sorted_cats:
        if cat.Id in selected_cat_ids:
            category_names.append(cat.Name)
    
    legend_view, error_msg = create_legend_view(
        doc, category_names, param_name, values_list, colors_dict, solid_fill
    )
    
    if error_msg:
        forms.alert(u'Ошибка создания легенды:\n{0}'.format(error_msg))
    # Успех - не показываем сообщение

elif action == 'material_legend':
    # Получаем имена выбранных категорий
    category_names = []
    for cat in form.sorted_cats:
        if cat.Id in selected_cat_ids:
            category_names.append(cat.Name)
    
    # Собираем материалы с элементов
    materials_list = collect_materials_from_elements(doc, elements_on_view, selected_cat_ids)
    
    if not materials_list:
        forms.alert(u'Не найдено материалов на элементах выбранных категорий.')
    else:
        legend_view, error_msg = create_material_legend_view(
            doc, category_names, materials_list, solid_fill
        )
        
        if error_msg:
            forms.alert(u'Ошибка создания легенды материалов:\n{0}'.format(error_msg))
        # Успех - не показываем сообщение
