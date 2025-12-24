# -*- coding: utf-8 -*-
# PyRevit script: автоматическая проставка высотных отметок Spot Elevation
# Работает в активном виде (план, 3D, разрез и т.п.)
# Пользователь выбирает трубы, воздуховоды, кабельные лотки.
# Для вертикальных и наклонных труб ставятся отметки на обоих концах.
# Для горизонтальных труб — одна отметка на конце.
# Для воздуховодов и кабельных лотков — отметка по низу.
# Тип высотной отметки берется жестко: ADSK_Схема_Проектная_Отметка снизу_Вверх
# Выноска горизонтальная, длина ~10 мм на листе.
# Сторона (влево/вправо) выбирается по соседним МЭП-элементам на виде:
# выноска ставится туда, где рядом меньше соседних элементов (без анализа аннотаций).
#
# Поддержка "Переместить элементы" (Displace Elements) на 3D виде:
# Если элементы смещены для разнесённого вида, отметки ставятся с учётом смещения,
# но значение отметки берётся от реальной позиции элемента.

from Autodesk.Revit.DB import (
    XYZ,
    FilteredElementCollector,
    SpotDimensionType,
    BuiltInParameter,
    Reference,
    FamilyInstance,
    FamilySymbol,
    BoundingBoxXYZ,
    BuiltInCategory,
    ElementId,
    DisplacementElement,
    View3D,
    AnnotationSymbol,
    IndependentTag,
    TagMode,
    TagOrientation,
    Options,
    Solid,
    GeometryInstance,
)
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray
from Autodesk.Revit.UI import Selection
from System.Collections.Generic import List
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

TOL = 1e-6
SPOT_TYPE_NAME = u"ADSK_Схема_Проектная_Отметка снизу_Вверх"
PAPER_LEADER_MM = 10.0  # длина выноски на листе, мм

# Имена марок для смещённых элементов (DisplacementElement)
TAG_PIPE_NAME = u"ADSK_M_Трубы_Высотная отметка"
TAG_DUCT_NAME = u"ADSK_M_Воздуховоды_Высотная отметка"
TAG_CABLETRAY_NAME = None  # Для кабельных лотков пока не указано

# Категории, которые считаем МЭП-соседями (geometry only, без аннотаций)
MEP_NEIGHBOR_CAT_IDS = set([
    int(BuiltInCategory.OST_PipeCurves),
    int(BuiltInCategory.OST_DuctCurves),
    int(BuiltInCategory.OST_CableTray),
])

# Кэш для смещений элементов на 3D виде (DisplacementElement)
displacement_cache = {}  # element_id -> XYZ displacement
displacement_elements_cache = {}  # displacement_element_id -> XYZ displacement
displacement_error = None


def build_displacement_cache():
    """Построить словарь смещений для DisplacementElement на активном 3D виде."""
    global displacement_cache, displacement_elements_cache, displacement_error
    displacement_cache = {}
    displacement_elements_cache = {}
    displacement_error = None
    
    # Работает только на 3D видах
    if not isinstance(active_view, View3D):
        return
    
    view_id = active_view.Id
    found_count = 0
    debug_info = []
    
    try:
        # Ищем все DisplacementElement в документе
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        for elem in all_elements:
            disp_elem = None
            try:
                if elem.GetType().Name == "DisplacementElement":
                    disp_elem = elem
            except:
                pass
            
            if disp_elem is None:
                continue
            
            found_count += 1
            disp_id = disp_elem.Id.IntegerValue
            
            try:
                # Проверяем, принадлежит ли этот DisplacementElement нашему виду
                try:
                    owner_view_id = disp_elem.OwnerViewId
                    if owner_view_id.IntegerValue != view_id.IntegerValue:
                        continue
                except:
                    pass  # Если не удалось проверить - берём всё равно
                
                # Получаем смещение из параметров
                disp_x = 0.0
                disp_y = 0.0
                disp_z = 0.0
                
                for p in disp_elem.Parameters:
                    try:
                        pname = p.Definition.Name
                        if pname in [u"Смещение X", "Displacement X"]:
                            disp_x = p.AsDouble()
                        elif pname in [u"Смещение Y", "Displacement Y"]:
                            disp_y = p.AsDouble()
                        elif pname in [u"Смещение Z", "Displacement Z"]:
                            disp_z = p.AsDouble()
                    except:
                        continue
                
                if abs(disp_x) < TOL and abs(disp_y) < TOL and abs(disp_z) < TOL:
                    continue
                
                total_displacement = XYZ(disp_x, disp_y, disp_z)
                displacement_elements_cache[disp_id] = total_displacement
                
                debug_info.append(u"DE#{0}: ({1:.1f}, {2:.1f}, {3:.1f})m".format(
                    disp_id, disp_x * 0.3048, disp_y * 0.3048, disp_z * 0.3048))
                        
            except Exception as inner_ex:
                debug_info.append(u"DE#{0}: error={1}".format(disp_id, str(inner_ex)))
                continue
        
        if found_count == 0:
            displacement_error = u"DisplacementElement не найдены"
        else:
            displacement_error = u"Найдено DE: {0}, с смещением: {1}\n{2}".format(
                found_count, len(displacement_elements_cache), "\n".join(debug_info))
                
    except Exception as ex:
        displacement_error = str(ex)


def find_element_displacement(element):
    """
    Найти смещение для конкретного элемента, сравнивая его BoundingBox 
    на виде с реальным положением.
    """
    if not displacement_elements_cache:
        return XYZ(0, 0, 0)
    
    # Получаем реальное положение элемента (из Location)
    real_pos = None
    try:
        loc = element.Location
        if loc:
            curve = getattr(loc, 'Curve', None)
            if curve:
                real_pos = curve.GetEndPoint(1)  # Конец элемента
            else:
                point = getattr(loc, 'Point', None)
                if point:
                    real_pos = point
    except:
        pass
    
    if real_pos is None:
        return XYZ(0, 0, 0)
    
    # Получаем BoundingBox на виде (может быть смещён)
    try:
        bb = element.get_BoundingBox(active_view)
        if bb and bb.Max:
            view_pos = bb.Max  # Используем Max как приблизительную позицию
            
            # Вычисляем разницу (смещение)
            diff = view_pos - real_pos
            
            # Проверяем, совпадает ли это смещение с каким-то DisplacementElement
            for de_id, de_disp in displacement_elements_cache.items():
                # Сравниваем с погрешностью
                if (abs(diff.X - de_disp.X) < 0.1 and 
                    abs(diff.Y - de_disp.Y) < 0.1 and 
                    abs(diff.Z - de_disp.Z) < 0.1):
                    return de_disp
            
            # Если точного совпадения нет, но есть значительное смещение - 
            # берём ближайшее из имеющихся
            if diff.GetLength() > 1.0:  # Больше ~30 см
                best_match = None
                best_dist = float('inf')
                for de_id, de_disp in displacement_elements_cache.items():
                    dist = (diff - de_disp).GetLength()
                    if dist < best_dist:
                        best_dist = dist
                        best_match = de_disp
                if best_match and best_dist < 10:  # В пределах ~3м погрешности
                    return best_match
    except:
        pass
    
    return XYZ(0, 0, 0)


def get_element_displacement(element_id):
    """Получить вектор смещения для элемента (или нулевой вектор, если нет смещения)."""
    if element_id.IntegerValue in displacement_cache:
        return displacement_cache[element_id.IntegerValue]
    return XYZ(0, 0, 0)


class DisplacedElementSelectionFilter(Selection.ISelectionFilter):
    """
    Фильтр для выбора элементов, включая смещённые через DisplacementElement.
    Разрешает выбор труб, воздуховодов, кабельных лотков.
    Для выбора отдельных элементов внутри набора смещения - 
    используйте Tab перед запуском скрипта или предвыберите элементы.
    """
    def AllowElement(self, element):
        try:
            # Разрешаем выбор DisplacementElement (для выбора всего набора)
            if isinstance(element, DisplacementElement):
                return True
            if isinstance(element, Pipe):
                return True
            if isinstance(element, Duct):
                return True
            if isinstance(element, CableTray):
                return True
        except:
            pass
        return False

    def AllowReference(self, reference, point):
        return True  # Разрешаем все ссылки для поддержки подэлементов


def get_element_from_reference(ref):
    """
    Получить элемент из Reference.
    Если Reference указывает на DisplacementElement, пытаемся получить вложенный элемент.
    """
    try:
        el = doc.GetElement(ref.ElementId)
        if el is None:
            return None
        
        # Если это DisplacementElement - пробуем получить LinkedElementId
        if isinstance(el, DisplacementElement):
            try:
                linked_id = ref.LinkedElementId
                if linked_id and linked_id.IntegerValue > 0:
                    linked_el = doc.GetElement(linked_id)
                    if linked_el and isinstance(linked_el, (Pipe, Duct, CableTray)):
                        return linked_el
            except:
                pass
        
        return el
    except:
        return None


def extract_mep_from_selection(selected_elements, selected_ids):
    """
    Извлекает МЕП-элементы из выбора, включая элементы внутри DisplacementElement.
    """
    elements = []
    
    for el in selected_elements:
        if el is None:
            continue
        
        eid_int = el.Id.IntegerValue
        
        # Если это DisplacementElement - извлекаем из него МЕП-элементы
        if isinstance(el, DisplacementElement):
            try:
                displaced_ids = el.GetDisplacedElementIds()
                if displaced_ids:
                    for displaced_id in displaced_ids:
                        if displaced_id.IntegerValue in selected_ids:
                            continue
                        displaced_el = doc.GetElement(displaced_id)
                        if displaced_el is None:
                            continue
                        if isinstance(displaced_el, Pipe) or isinstance(displaced_el, Duct) or isinstance(displaced_el, CableTray):
                            elements.append(displaced_el)
                            selected_ids.add(displaced_id.IntegerValue)
            except:
                pass
        elif isinstance(el, Pipe) or isinstance(el, Duct) or isinstance(el, CableTray):
            if eid_int not in selected_ids:
                elements.append(el)
                selected_ids.add(eid_int)
    
    return elements


def get_selected_elements():
    sel_ids = list(uidoc.Selection.GetElementIds())
    elements = []
    selected_ids = set()  # Для отслеживания уже выбранных
    
    # Сначала добавляем уже выделенные элементы (включая из DisplacementElement)
    preselected = [doc.GetElement(eid) for eid in sel_ids]
    elements = extract_mep_from_selection(preselected, selected_ids)

    # Если ничего не выбрано - просим выбрать
    if not elements:
        try:
            refs = uidoc.Selection.PickObjects(
                Selection.ObjectType.Element,
                DisplacedElementSelectionFilter(),
                u"Выберите элементы. Для выбора внутри набора смещения: Tab до запуска скрипта. Enter/Esc для завершения"
            )
        except:
            return []
        
        # Используем get_element_from_reference для получения элементов из ссылок
        picked_elements = [get_element_from_reference(r) for r in refs]
        elements = extract_mep_from_selection(picked_elements, selected_ids)
    else:
        # Если уже есть выбранные - спрашиваем, хочет ли пользователь добавить ещё
        add_more = forms.alert(
            u"Выбрано элементов: {0}\n\nДобавить ещё элементов к выбору?\n\n(Для выбора внутри набора смещения используйте Tab до запуска скрипта)".format(len(elements)),
            title=u"Высотные отметки",
            yes=True,
            no=True
        )
        if add_more:
            try:
                refs = uidoc.Selection.PickObjects(
                    Selection.ObjectType.Element,
                    DisplacedElementSelectionFilter(),
                    u"Выберите дополнительные элементы. Enter/Esc для завершения"
                )
                # Используем get_element_from_reference для получения элементов из ссылок
                picked_elements = [get_element_from_reference(r) for r in refs]
                additional = extract_mep_from_selection(picked_elements, selected_ids)
                elements.extend(additional)
            except:
                pass  # Пользователь отменил - работаем с тем что есть
    
    return elements


def build_neighbor_points():
    """Собираем центры соседних МЭП-элементов на активном виде (для выбора стороны выноски)."""
    neighbors = []
    if active_view is None:
        return neighbors

    col = FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType().ToElements()

    for el in col:
        cat = el.Category
        if not cat:
            continue
        if cat.Id.IntegerValue not in MEP_NEIGHBOR_CAT_IDS:
            continue
        try:
            bb = el.get_BoundingBox(active_view)
        except:
            bb = None
        if not isinstance(bb, BoundingBoxXYZ) or bb.Min is None or bb.Max is None:
            continue
        center = (bb.Min + bb.Max) * 0.5
        neighbors.append((el.Id, center))
    return neighbors


def get_spot_type_fixed():
    types = list(FilteredElementCollector(doc).OfClass(SpotDimensionType).ToElements())
    for stype in types:
        try:
            name_param = stype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if name_param and name_param.AsString() == SPOT_TYPE_NAME:
                return stype
        except:
            continue

    forms.alert(
        u"Не найден тип высотной отметки с именем '{0}'.\nПроверь, что он загружен в проект и имя совпадает.".format(
            SPOT_TYPE_NAME),
        exitscript=True
    )
    return None


def ask_axis_or_bottom_for_pipes():
    options = [
        u"По оси трубы",
        u"По низу трубы (приблизительно)"
    ]

    chosen = forms.SelectFromList.show(
        options,
        title=u"Откуда брать отметку относительно трубы?",
        multiselect=False
    )

    if not chosen:
        return None

    if u"оси" in chosen:
        return "axis"
    return "bottom"


def get_bottom_face_reference(element, target_point):
    """
    Получить Reference на нижнюю грань элемента вблизи target_point.
    Возвращает (reference, point_on_face) или (None, None).
    """
    try:
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = False

        geom = element.get_Geometry(opt)
        if geom is None:
            return None, None

        # Используем словарь для хранения результата (вместо nonlocal)
        result = {'best_ref': None, 'best_point': None, 'best_z': float('inf')}

        def process_solid(solid):
            if solid is None:
                return

            faces = solid.Faces
            for face in faces:
                try:
                    # Получаем BoundingBox грани
                    bb = face.GetBoundingBox()
                    if bb is None:
                        continue

                    # Получаем UV-параметры центра грани
                    uv_min = bb.Min
                    uv_max = bb.Max
                    uv_mid = (uv_min + uv_max) * 0.5

                    # Получаем 3D точку на грани
                    pt = face.Evaluate(uv_mid)
                    if pt is None:
                        continue

                    # Проверяем, что точка близка к target_point по X,Y
                    dist_xy = ((pt.X - target_point.X)**2 + (pt.Y - target_point.Y)**2)**0.5
                    if dist_xy > 1.0:  # ~30 см допуск
                        continue

                    # Ищем грань с минимальной Z (самая нижняя)
                    if pt.Z < result['best_z']:
                        ref = face.Reference
                        if ref is not None:
                            result['best_z'] = pt.Z
                            result['best_ref'] = ref
                            result['best_point'] = pt
                except:
                    continue

        for geom_obj in geom:
            if isinstance(geom_obj, Solid):
                process_solid(geom_obj)
            elif isinstance(geom_obj, GeometryInstance):
                inst_geom = geom_obj.GetInstanceGeometry()
                if inst_geom:
                    for inst_obj in inst_geom:
                        if isinstance(inst_obj, Solid):
                            process_solid(inst_obj)

        return result['best_ref'], result['best_point']
    except:
        return None, None


def get_pipe_points(pipe, mode):
    """
    Возвращает список кортежей (displaced_point, real_z) для трубы.
    displaced_point - координаты с учётом смещения (DisplacementElement)
    real_z - реальная Z-координата для значения отметки
    """
    loc = pipe.Location
    if not loc:
        return []

    curve = getattr(loc, 'Curve', None)
    if not curve:
        return []

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    # Сохраняем оригинальные Z-координаты (для значения отметки)
    original_z0 = p0.Z
    original_z1 = p1.Z

    c0 = XYZ(p0.X, p0.Y, p0.Z)
    c1 = XYZ(p1.X, p1.Y, p1.Z)

    if mode == 'bottom':
        # Получаем ВНЕШНИЙ диаметр трубы (наружный)
        radius = 0.0

        # Способ 1: Внешний диаметр
        if radius <= TOL:
            try:
                outer_diam_param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                if outer_diam_param:
                    val = outer_diam_param.AsDouble()
                    if val > TOL:
                        radius = val / 2.0
            except:
                pass

        # Способ 2: Обычный диаметр
        if radius <= TOL:
            try:
                diam_param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                if diam_param:
                    val = diam_param.AsDouble()
                    if val > TOL:
                        radius = val / 2.0
            except:
                pass

        # Способ 3: Свойство Diameter
        if radius <= TOL:
            try:
                val = pipe.Diameter
                if val > TOL:
                    radius = val / 2.0
            except:
                pass

        # Способ 4: Диаметр из коннекторов
        if radius <= TOL:
            try:
                conn_mgr = pipe.ConnectorManager
                if conn_mgr:
                    connectors = conn_mgr.Connectors
                    for conn in connectors:
                        if conn.Shape == 0:  # 0 = Round
                            val = conn.Radius
                            if val > TOL:
                                radius = val
                                break
            except:
                pass

        # Смещаем вниз на радиус (для позиционирования отметки)
        c0 = XYZ(c0.X, c0.Y, c0.Z - radius)
        c1 = XYZ(c1.X, c1.Y, c1.Z - radius)

        # Корректируем значения высоты: вычитаем радиус из оригинальных Z
        original_z0 = original_z0 - radius
        original_z1 = original_z1 - radius

    d = p1 - p0
    is_horizontal = abs(d.Z) < TOL

    # Длина трубы в футах
    pipe_length = d.GetLength()
    # 1 метр = 3.28084 фута
    ONE_METER_FT = 3.28084

    # Получаем вектор смещения для этого элемента
    displacement = get_element_displacement(pipe.Id)

    # Применяем смещение к координатам (для позиционирования отметки)
    # но сохраняем реальную Z-координату для значения отметки
    c0_displaced = c0 + displacement
    c1_displaced = c1 + displacement

    # Возвращаем кортежи (displaced_point, real_z)
    # real_z теперь содержит высоту низа трубы (если mode='bottom')
    if is_horizontal:
        return [(c1_displaced, original_z1)]
    else:
        # Для наклонных/вертикальных труб короче 1 метра - только одна отметка в конце
        if pipe_length < ONE_METER_FT:
            return [(c1_displaced, original_z1)]
        return [(c0_displaced, original_z0), (c1_displaced, original_z1)]


def get_duct_points(duct):
    """
    Возвращает список кортежей для воздуховодов.
    Для круглых: (displaced_point, real_z)
    Для прямоугольных: (displaced_point, perp, edge_offset, real_z)
    """
    # Определяем, круглый или прямоугольный воздуховод
    diam_param = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    diam = 0.0
    if diam_param:
        try:
            diam = diam_param.AsDouble()
        except:
            diam = 0.0

    height_param = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
    h = 0.0
    if height_param:
        try:
            h = height_param.AsDouble()
        except:
            h = 0.0

    width_param = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    w = 0.0
    if width_param:
        try:
            w = width_param.AsDouble()
        except:
            w = 0.0

    loc = duct.Location
    if not loc:
        return []

    curve = getattr(loc, 'Curve', None)
    if not curve:
        return []

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    c0 = XYZ(p0.X, p0.Y, p0.Z)
    c1 = XYZ(p1.X, p1.Y, p1.Z)

    # Получаем вектор смещения для этого элемента
    displacement = get_element_displacement(duct.Id)

    is_round = diam > TOL

    if is_round:
        # Круглый воздуховод — отметка по НИЗУ (смещаем на радиус вниз)
        radius = diam / 2.0
        c0_bottom = XYZ(c0.X, c0.Y, c0.Z - radius)
        c1_bottom = XYZ(c1.X, c1.Y, c1.Z - radius)
        
        # Применяем смещение
        c0_bottom_displaced = c0_bottom + displacement
        c1_bottom_displaced = c1_bottom + displacement
        
        d = p1 - p0
        is_horizontal = abs(d.Z) < TOL
        if is_horizontal:
            return [(c1_bottom_displaced, c1_bottom.Z)]
        else:
            # Для наклонных/вертикальных - проверяем длину (как для труб)
            pipe_length = d.GetLength()
            ONE_METER_FT = 3.28084
            if pipe_length < ONE_METER_FT:
                return [(c1_bottom_displaced, c1_bottom.Z)]
            return [(c0_bottom_displaced, c0_bottom.Z), (c1_bottom_displaced, c1_bottom.Z)]
    else:
        # Прямоугольный — по НИЗУ
        # Вычисляем направление перпендикулярно оси воздуховода (в горизонтальной плоскости)
        duct_dir = p1 - p0
        duct_dir_horiz = XYZ(duct_dir.X, duct_dir.Y, 0)
        if duct_dir_horiz.GetLength() > TOL:
            duct_dir_horiz = duct_dir_horiz.Normalize()
            # Перпендикуляр в горизонтальной плоскости (поворот на 90 градусов)
            perp = XYZ(-duct_dir_horiz.Y, duct_dir_horiz.X, 0)
        else:
            # Вертикальный воздуховод - берём произвольное направление
            perp = XYZ(1, 0, 0)
        
        # Смещение от центра к краю на половину ширины
        edge_offset = w / 2.0 if w > 0 else 0.0
        
        # Точка: конец воздуховода, низ по Z
        z_bottom = c1.Z - h / 2.0 if h > 0 else c1.Z
        c1_bottom = XYZ(c1.X, c1.Y, z_bottom)
        
        # Применяем смещение
        c1_bottom_displaced = c1_bottom + displacement
        
        # Возвращаем точку и информацию о смещении (перпендикуляр, полуширина, реальная Z)
        return [(c1_bottom_displaced, perp, edge_offset, c1_bottom.Z)]


def get_cable_tray_points(cable_tray):
    """
    Получить точки для кабельного лотка (отметка по низу).
    Возвращает список кортежей: (displaced_point, perp, edge_offset, real_z)
    """
    loc = cable_tray.Location
    if not loc:
        return []

    curve = getattr(loc, 'Curve', None)
    if not curve:
        return []

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    # Получаем высоту лотка
    height = 0.0
    try:
        h_param = cable_tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
        if h_param and h_param.HasValue:
            height = h_param.AsDouble()
    except:
        pass

    # Получаем ширину лотка
    width = 0.0
    try:
        w_param = cable_tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        if w_param and w_param.HasValue:
            width = w_param.AsDouble()
    except:
        pass

    # Смещение на низ (height/2)
    z_offset = height / 2.0 if height > 0 else 0.0
    
    c1 = XYZ(p1.X, p1.Y, p1.Z)
    c1_bottom = XYZ(c1.X, c1.Y, c1.Z - z_offset)

    # Получаем вектор смещения для этого элемента
    displacement = get_element_displacement(cable_tray.Id)
    c1_bottom_displaced = c1_bottom + displacement

    # Вычисляем перпендикуляр для смещения к краю
    tray_dir = p1 - p0
    tray_dir_horiz = XYZ(tray_dir.X, tray_dir.Y, 0)
    if tray_dir_horiz.GetLength() > TOL:
        tray_dir_horiz = tray_dir_horiz.Normalize()
        perp = XYZ(-tray_dir_horiz.Y, tray_dir_horiz.X, 0)
    else:
        perp = XYZ(1, 0, 0)

    edge_offset = width / 2.0 if width > 0 else 0.0

    # Возвращаем кортеж с реальной Z-координатой
    return [(c1_bottom_displaced, perp, edge_offset, c1_bottom.Z)]


def project_point_to_view_plane(point, view):
    # Для SpotElevation достаточно использовать мировые координаты
    return point


def choose_side_by_neighbors(view, origin, element_id, neighbors, offset_ft):
    """Определяем, куда ставить выноску (влево/вправо) по соседним элементам."""
    try:
        right = view.RightDirection
        up = view.UpDirection
    except:
        right = XYZ(1, 0, 0)
        up = XYZ(0, 1, 0)

    if right.GetLength() < TOL:
        right = XYZ(1, 0, 0)
    if up.GetLength() < TOL:
        up = XYZ(0, 1, 0)

    right = right.Normalize()
    up = up.Normalize()

    # Радиус поиска соседей — несколько длин выноски
    search_half = offset_ft * 4.0

    left_count = 0
    right_count = 0

    for nid, cpt in neighbors:
        if nid == element_id:
            continue
        v = cpt - origin
        dx = v.DotProduct(right)
        dy = v.DotProduct(up)

        if abs(dx) > search_half or abs(dy) > search_half:
            continue

        if dx > 0:
            right_count += 1
        elif dx < 0:
            left_count += 1

    # Сторона с меньшим количеством соседей
    if left_count < right_count:
        side_sign = -1.0  # влево
    elif right_count < left_count:
        side_sign = 1.0   # вправо
    else:
        side_sign = 1.0   # по умолчанию вправо

    return right, side_sign


def get_leader_points(view, origin, element_id, neighbors, extra_offset_vec=None):
    # Горизонтальная выноска длиной ~10 мм на листе.
    # extra_offset_vec - дополнительное смещение (напр. от края воздуховода)
    try:
        scale = view.Scale
    except:
        scale = 100

    offset_m = (PAPER_LEADER_MM / 1000.0) * float(scale)
    offset_ft = offset_m / 0.3048

    right, side_sign = choose_side_by_neighbors(view, origin, element_id, neighbors, offset_ft)
    right_norm = right.Normalize()
    
    # Базовое смещение выноски (всегда горизонтально на виде)
    offset_vec = right_norm.Multiply(offset_ft * side_sign)
    
    # Добавляем дополнительное смещение (от края элемента) в ту же сторону
    # Проецируем extra_offset_vec на горизонтальное направление вида
    if extra_offset_vec is not None:
        # Проекция на горизонталь вида (right_norm)
        proj_len = extra_offset_vec.DotProduct(right_norm)
        if abs(proj_len) > TOL:
            # Берём только горизонтальную составляющую смещения
            extra_horiz = right_norm.Multiply(abs(proj_len) * side_sign)
            offset_vec = offset_vec + extra_horiz

    bend = origin + offset_vec
    
    # end - конечная точка выноски (где текст)
    # Смещаем дальше в ту же сторону, чтобы текст был ориентирован правильно
    text_offset = right_norm.Multiply(offset_ft * 0.5 * side_sign)  # Небольшое смещение для ориентации текста
    end = bend + text_offset
    
    return bend, end, side_sign


def get_annotation_symbol_by_name(family_name):
    """Найти типоразмер аннотационного семейства по имени семейства."""
    if not family_name:
        return None
    
    # Ищем среди всех FamilySymbol
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    
    for symbol in collector:
        try:
            # Проверяем имя семейства
            family = symbol.Family
            if family and family.Name == family_name:
                # Активируем, если не активен
                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()
                return symbol
        except:
            continue
    
    return None


def get_tag_symbol_for_element(element):
    """Получить типоразмер марки для элемента в зависимости от его типа."""
    if isinstance(element, Pipe):
        return get_annotation_symbol_by_name(TAG_PIPE_NAME)
    elif isinstance(element, Duct):
        return get_annotation_symbol_by_name(TAG_DUCT_NAME)
    elif isinstance(element, CableTray):
        if TAG_CABLETRAY_NAME:
            return get_annotation_symbol_by_name(TAG_CABLETRAY_NAME)
    return None


def create_elevation_annotation(view, element, point, neighbors, failed_elements=None):
    """
    Создать марку высотной отметки для смещённого элемента.
    Марка размещается с выноской от точки point (смещённая позиция).
    """
    tag_symbol = get_tag_symbol_for_element(element)
    
    if tag_symbol is None:
        elem_type = u"неизвестен"
        if isinstance(element, Pipe):
            elem_type = u"труба"
        elif isinstance(element, Duct):
            elem_type = u"воздуховод"
        elif isinstance(element, CableTray):
            elem_type = u"лоток"
        if failed_elements is not None:
            failed_elements.append((element, u"Не найдена марка для типа: {}".format(elem_type)))
        return None
    
    try:
        # Создаём Reference на элемент
        ref = Reference(element)
        
        # Вычисляем позицию для головки марки (со смещением как для SpotElevation)
        bend, end, side_sign = get_leader_points(view, point, element.Id, neighbors, None)
        
        # Создаём марку через IndependentTag.Create
        # Марка привязывается к элементу, размещаем головку со смещением
        tag = IndependentTag.Create(
            doc,
            view.Id,
            ref,
            True,  # addLeader - С ВЫНОСКОЙ
            TagMode.TM_ADDBY_CATEGORY,  # Режим маркировки по категории
            TagOrientation.Horizontal,  # Горизонтальная ориентация
            end  # Точка размещения головки марки (конец выноски)
        )
        
        # Меняем тип марки на нужный
        if tag and tag_symbol:
            try:
                tag.ChangeTypeId(tag_symbol.Id)
            except:
                pass
        
        # Устанавливаем позицию головки марки
        if tag:
            try:
                tag.TagHeadPosition = end
            except:
                pass
        
        return tag
    except Exception as e:
        if failed_elements is not None:
            failed_elements.append((element, u"Ошибка создания марки: {}".format(str(e))))
        return None


def create_spot_elevation(view, spot_type, element, point, neighbors, failed_elements=None, extra_offset_info=None, use_bottom_ref=False, real_z=None, is_displaced=False):
    # Создаём высотную отметку.
    # point - смещённая позиция (где визуально находится элемент)
    # real_z - реальная Z-координата для значения отметки
    # extra_offset_info = (perp_direction, offset_distance) - доп. смещение для прямоугольных воздуховодов
    # use_bottom_ref - использовать Reference на нижнюю грань (для режима "по низу")
    # is_displaced - элемент смещён через DisplacementElement

    base_origin = project_point_to_view_plane(point, view)

    # Для прямоугольных воздуховодов: смещаем origin к краю
    origin = base_origin
    if extra_offset_info is not None:
        perp, offset_dist = extra_offset_info
        if offset_dist > 0:
            # Сначала определяем сторону выноски
            bend_test, end_test, side_sign = get_leader_points(view, base_origin, element.Id, neighbors, None)

            # Смещаем origin к краю в сторону выноски
            dot = perp.DotProduct(view.RightDirection)
            if dot * side_sign < 0:
                perp = perp.Multiply(-1)
            origin = XYZ(base_origin.X + perp.X * offset_dist,
                         base_origin.Y + perp.Y * offset_dist,
                         base_origin.Z)

    bend, end, side_sign = get_leader_points(view, origin, element.Id, neighbors, None)

    # Создаём Reference
    reference = None
    bottom_point = None

    if use_bottom_ref and isinstance(element, Pipe):
        # Для труб в режиме "по низу" - пробуем получить Reference на нижнюю грань
        bottom_ref, bottom_point = get_bottom_face_reference(element, point)
        if bottom_ref is not None:
            reference = bottom_ref

    # Если не удалось получить Reference на грань - используем Reference на элемент
    if reference is None:
        try:
            reference = Reference(element)
        except Exception as e:
            if failed_elements is not None:
                failed_elements.append((element, u"Не удалось создать Reference: {}".format(str(e))))
            return None

    # ref_pt - точка для расчёта высоты
    # Если получили точку на нижней грани - используем её
    if bottom_point is not None:
        ref_pt = bottom_point
    elif real_z is not None:
        ref_pt = XYZ(base_origin.X, base_origin.Y, real_z)
    else:
        ref_pt = base_origin

    spot = None
    
    # Для смещённых элементов - создаём БЕЗ выноски, иначе выноска пойдёт к реальной геометрии
    if is_displaced:
        try:
            # Без выноски - только отметка в точке origin
            spot = doc.Create.NewSpotElevation(
                view,
                reference,
                origin,  # point - где разместить
                origin,  # bend = origin (нет выноски)
                origin,  # end = origin (нет выноски)
                ref_pt,  # точка отсчёта высоты
                False    # hasLeader = False
            )
        except Exception as e:
            if failed_elements is not None:
                failed_elements.append((element, u"Ошибка (displaced): {}".format(str(e))))
            return None
    else:
        # Для обычных элементов - с выноской
        try:
            spot = doc.Create.NewSpotElevation(
                view,
                reference,
                origin,
                bend,
                end,
                ref_pt,
                True  # hasLeader
            )
        except Exception as e1:
            # fallback: без выноски
            try:
                spot = doc.Create.NewSpotElevation(
                    view,
                    reference,
                    origin,
                    origin,
                    origin,
                    ref_pt,
                    False
                )
            except Exception as e2:
                if failed_elements is not None:
                    cat_name = u"?"
                    try:
                        cat_name = element.Category.Name
                    except:
                        pass
                    failed_elements.append((element, u"Категория: {}, Ошибка: {}".format(cat_name, str(e2))))
                return None

    try:
        spot.ChangeTypeId(spot_type.Id)
    except:
        pass

    return spot


def main():
    if active_view is None:
        forms.alert(u"Нет активного вида.", exitscript=True)
        return

    elements = get_selected_elements()
    if not elements:
        forms.alert(u"Не выбрано ни одного подходящего элемента.", exitscript=True)
        return

    spot_type = get_spot_type_fixed()
    if spot_type is None:
        return

    has_pipes = any(isinstance(el, Pipe) for el in elements)
    mode = "axis"
    if has_pipes:
        m = ask_axis_or_bottom_for_pipes()
        if m is None:
            return
        mode = m

    # Строим кэш смещений для 3D вида (DisplacementElement)
    build_displacement_cache()
    
    # Для каждого выбранного элемента определяем его смещение
    for el in elements:
        disp = find_element_displacement(el)
        if disp.GetLength() > TOL:
            displacement_cache[el.Id.IntegerValue] = disp
    
    # Подсчитываем, сколько выбранных элементов смещены
    displaced_count = sum(1 for el in elements if el.Id.IntegerValue in displacement_cache)

    # Список соседей для всех элементов на виде
    neighbors = build_neighbor_points()

    created_count = 0
    created_ids = []
    failed_elements = []
    skipped_no_point = 0

    with revit.Transaction(u"Высотные отметки Spot Elevation"):
        for el in elements:
            pts_data = []  # Список кортежей (point, extra_offset_info, use_bottom_ref, real_z)
            use_bottom = (mode == 'bottom')  # Для труб в режиме "низ"

            if isinstance(el, Pipe):
                # Для труб: возвращает кортежи (displaced_point, real_z)
                pipe_pts = get_pipe_points(el, mode)
                for item in pipe_pts:
                    if isinstance(item, tuple) and len(item) == 2:
                        pt, real_z = item
                        pts_data.append((pt, None, use_bottom, real_z))
                    else:
                        # Обратная совместимость (если старый формат)
                        pts_data.append((item, None, use_bottom, None))
            elif isinstance(el, Duct):
                # Для воздуховодов:
                # Круглые: (displaced_point, real_z)
                # Прямоугольные: (displaced_point, perp, offset, real_z)
                duct_result = get_duct_points(el)
                for item in duct_result:
                    if isinstance(item, tuple) and len(item) == 4:
                        # Прямоугольный воздуховод: (point, perp, offset, real_z)
                        pt, perp, offset, real_z = item
                        pts_data.append((pt, (perp, offset), True, real_z))
                    elif isinstance(item, tuple) and len(item) == 2:
                        # Круглый воздуховод: (displaced_point, real_z)
                        pt, real_z = item
                        pts_data.append((pt, None, True, real_z))
                    else:
                        # Обратная совместимость
                        pts_data.append((item, None, True, None))
            elif isinstance(el, CableTray):
                # Для кабельных лотков: (displaced_point, perp, offset, real_z)
                tray_result = get_cable_tray_points(el)
                for item in tray_result:
                    if isinstance(item, tuple) and len(item) == 4:
                        pt, perp, offset, real_z = item
                        pts_data.append((pt, (perp, offset), True, real_z))
                    elif isinstance(item, tuple) and len(item) == 3:
                        # Старый формат без real_z
                        pt, perp, offset = item
                        pts_data.append((pt, (perp, offset), True, None))
                    else:
                        pts_data.append((item, None, True, None))

            for pt_info in pts_data:
                if pt_info is None:
                    continue
                pt, extra_offset, use_bottom_ref, real_z = pt_info
                if pt is None:
                    continue
                # Проверяем, смещён ли элемент
                is_displaced = el.Id.IntegerValue in displacement_cache
                
                if is_displaced:
                    # Для смещённых элементов используем аннотационную марку с выноской
                    annotation = create_elevation_annotation(active_view, el, pt, neighbors, failed_elements)
                    if annotation:
                        created_count += 1
                        created_ids.append(annotation.Id)
                else:
                    # Для обычных элементов используем SpotElevation
                    spot = create_spot_elevation(active_view, spot_type, el, pt, neighbors, failed_elements, extra_offset, use_bottom_ref, real_z, is_displaced)
                    if spot:
                        created_count += 1
                        created_ids.append(spot.Id)

    # Выделяем созданные отметки, чтобы было проще их найти
    if created_ids:
        id_list = List[ElementId]()
        for sid in created_ids:
            id_list.Add(sid)
        uidoc.Selection.SetElementIds(id_list)

    # Формируем сообщение
    msg = u"Создано высотных отметок: {0}".format(created_count)
    if displaced_count > 0:
        msg += u"\nИз них для смещённых элементов (марки): {0}".format(displaced_count)
        msg += u"\n(для труб: {0}, для воздуховодов: {1})".format(TAG_PIPE_NAME, TAG_DUCT_NAME)
    if skipped_no_point > 0:
        msg += u"\nПропущено (не удалось определить точку): {0}".format(skipped_no_point)
    if failed_elements:
        msg += u"\nНе удалось создать отметку для {0} элементов:".format(len(failed_elements))
        # Показываем первые 3 ошибки
        for i, (el, reason) in enumerate(failed_elements[:3]):
            try:
                el_id = el.Id.IntegerValue
            except:
                el_id = "?"
            msg += u"\n  - ID {}: {}".format(el_id, reason)
        if len(failed_elements) > 3:
            msg += u"\n  ... и ещё {}".format(len(failed_elements) - 3)

    forms.alert(msg)


if __name__ == '__main__':
    main()