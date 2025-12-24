# -*- coding: utf-8 -*-
from __future__ import print_function, division

import sys
import clr
import re

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, ViewType, View3D, XYZ, Level,
    Transaction, ElementCategoryFilter, ConnectorType, ElementTransformUtils,
    UnitUtils, UnitTypeId, FamilyInstance, BuiltInParameter, DisplacementElement
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import revit, forms

uidoc = __revit__.ActiveUIDocument  # noqa
doc = uidoc.Document
view = doc.ActiveView

CATEGORIES = (
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_FlexPipeCurves,
)

FAMILY_NAME = u"ТипАн_Мрк_B4E_Уровень"

# Кэш для смещений элементов (DisplacementElement)
TOL = 1e-6
displacement_cache = {}  # element_id -> XYZ displacement
displacement_elements_cache = {}  # displacement_element_id -> XYZ displacement


def in_any_category(elem, cats):
    try:
        return elem.Category and elem.Category.Id.IntegerValue in [int(c) for c in cats]
    except:
        return False


def build_displacement_cache():
    """Построить словарь смещений для DisplacementElement на активном 3D виде."""
    global displacement_cache, displacement_elements_cache
    displacement_cache = {}
    displacement_elements_cache = {}
    
    if not isinstance(view, View3D):
        return
    
    view_id = view.Id
    
    try:
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
            
            disp_id = disp_elem.Id.IntegerValue
            
            try:
                try:
                    owner_view_id = disp_elem.OwnerViewId
                    if owner_view_id.IntegerValue != view_id.IntegerValue:
                        continue
                except:
                    pass
                
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
                        
            except:
                continue
                
    except:
        pass


def find_element_displacement(element):
    """Найти смещение для элемента, сравнивая BoundingBox на виде с реальным положением."""
    if not displacement_elements_cache:
        return XYZ(0, 0, 0)
    
    real_pos = None
    try:
        loc = element.Location
        if loc:
            curve = getattr(loc, 'Curve', None)
            if curve:
                real_pos = curve.GetEndPoint(1)
            else:
                point = getattr(loc, 'Point', None)
                if point:
                    real_pos = point
    except:
        pass
    
    if real_pos is None:
        return XYZ(0, 0, 0)
    
    try:
        bb = element.get_BoundingBox(view)
        if bb and bb.Max:
            view_pos = bb.Max
            diff = view_pos - real_pos
            
            for de_id, de_disp in displacement_elements_cache.items():
                if (abs(diff.X - de_disp.X) < 0.1 and 
                    abs(diff.Y - de_disp.Y) < 0.1 and 
                    abs(diff.Z - de_disp.Z) < 0.1):
                    return de_disp
            
            if diff.GetLength() > 1.0:
                best_match = None
                best_dist = float('inf')
                for de_id, de_disp in displacement_elements_cache.items():
                    dist = (diff - de_disp).GetLength()
                    if dist < best_dist:
                        best_dist = dist
                        best_match = de_disp
                if best_match and best_dist < 10:
                    return best_match
    except:
        pass
    
    return XYZ(0, 0, 0)


def get_element_displacement(element_id):
    """Получить вектор смещения для элемента."""
    if element_id.IntegerValue in displacement_cache:
        return displacement_cache[element_id.IntegerValue]
    return XYZ(0, 0, 0)


class VISElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        # Разрешаем выбор DisplacementElement
        if isinstance(element, DisplacementElement):
            return True
        return in_any_category(element, CATEGORIES)
    def AllowReference(self, reference, position):
        return True


def start_up_checks():
    if view is None or view.ViewType != ViewType.ThreeD:
        forms.alert(u"Добавление отметок возможно только на 3D-виде.", title=u"Ошибка", exitscript=True)
    if not getattr(view, "IsLocked", False):
        forms.alert(u"3D-вид должен быть заблокирован.", title=u"Ошибка", exitscript=True)


def extract_mep_from_selection(selected_elements, selected_ids):
    """Извлекает МЕП-элементы из выбора, включая элементы внутри DisplacementElement."""
    elements = []
    
    for el in selected_elements:
        if el is None:
            continue
        
        eid_int = el.Id.IntegerValue
        
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
                        if in_any_category(displaced_el, CATEGORIES):
                            elements.append(displaced_el)
                            selected_ids.add(displaced_id.IntegerValue)
            except:
                pass
        elif in_any_category(el, CATEGORIES):
            if eid_int not in selected_ids:
                elements.append(el)
                selected_ids.add(eid_int)
    
    return elements


def get_pre_selected():
    elems = [doc.GetElement(eid) for eid in uidoc.Selection.GetElementIds()]
    selected_ids = set()
    return extract_mep_from_selection(elems, selected_ids)


def get_selected():
    pre = get_pre_selected()
    if pre:
        return pre
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, VISElementsFilter(), 
            u"Выберите воздуховоды/трубы/арматуру/оборудование. Для выбора внутри набора смещения: Tab до запуска")
    except:
        sys.exit()
    picked_elements = [doc.GetElement(r.ElementId) for r in refs]
    selected_ids = set()
    return extract_mep_from_selection(picked_elements, selected_ids)


def get_name_safe(obj):
    # Пытаемся получить .Name разными способами (с учётом IronPython)
    try:
        return obj.Name
    except:
        pass
    try:
        return obj.get_Name()
    except:
        pass
    try:
        p = obj.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p: return p.AsString()
    except:
        pass
    return u""


def get_family_name_safe(sym):
    try:
        fam = sym.Family
        if fam:
            try:
                return fam.Name
            except:
                return fam.get_Name()
    except:
        pass
    # иногда семейство доступно только через параметр типа
    try:
        p = sym.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if p: return p.AsString() or u""
    except:
        pass
    return u""


def find_template_on_view():
    """Ищем Generic Annotation (FamilyInstance) с OwnerViewId == активному 3D-виду.
       Берём только семейство с именем FAMILY_NAME.
    """
    all_ga = [fi for fi in FilteredElementCollector(doc)              .OfCategory(BuiltInCategory.OST_GenericAnnotation)              .WhereElementIsNotElementType().ToElements()
              if isinstance(fi, FamilyInstance) and fi.OwnerViewId == view.Id]
    if not all_ga:
        forms.alert(u"На активном 3D-виде нет нужной аннотации.\n"
                    u"Разместите один экземпляр \"{0}\") и повторите.".format(FAMILY_NAME),
                    title=u"Ошибка", exitscript=True)
    # ищем только точное совпадение с FAMILY_NAME
    exact = []
    debug_info = []
    for fi in all_ga:
        try:
            sym = fi.Symbol
            sname = get_name_safe(sym)
            fname = get_family_name_safe(sym)
            debug_info.append(u"Тип: '{}', Семейство: '{}'".format(sname, fname))
            if fname == FAMILY_NAME or sname == FAMILY_NAME:
                exact.append(fi)
        except Exception as e:
            debug_info.append(u"Ошибка: {}".format(e))
    if not exact:
        msg = u"На активном 3D-виде нет семейства \"{0}\".\n\n".format(FAMILY_NAME)
        msg += u"Найденные аннотации:\n" + u"\n".join(debug_info)
        forms.alert(msg, title=u"Ошибка", exitscript=True)
    return exact[0]


def get_connectors(element):
    conns = []
    try:
        if hasattr(element, "MEPModel") and element.MEPModel and element.MEPModel.ConnectorManager:
            inst_cons = list(element.MEPModel.ConnectorManager.Connectors)
            filtered = [c for c in inst_cons if c.Domain.ToString() in ("DomainHvac", "DomainPiping")]
            if filtered:
                min_z = min(c.Origin.Z for c in filtered)
                max_z = max(c.Origin.Z for c in filtered)
                conns.extend([c for c in filtered if abs(c.Origin.Z - min_z) < 1e-6 or abs(c.Origin.Z - max_z) < 1e-6])
    except:
        pass

    try:
        if in_any_category(element, (BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves,
                                     BuiltInCategory.OST_FlexDuctCurves, BuiltInCategory.OST_FlexPipeCurves)):
            for c in element.ConnectorManager.Connectors:
                if c.ConnectorType != ConnectorType.Curve:
                    conns.append(c)
    except:
        pass

    return conns


def get_connector_coordinates(element):
    conns = get_connectors(element)
    start_point, end_point = None, None
    for c in conns:
        if c.Owner.Id == element.Id:
            if start_point is None:
                start_point = c.Origin
            else:
                end_point = c.Origin
                break

    if start_point is None and end_point is None:
        bb = element.get_BoundingBox(None)
        if not bb:
            return None, None
        start_point = XYZ((bb.Min.X + bb.Max.X)/2.0, (bb.Min.Y + bb.Max.Y)/2.0, bb.Min.Z)
        end_point   = XYZ(start_point.X, start_point.Y, bb.Max.Z)
        return start_point, end_point

    if end_point is None:
        bb = element.get_BoundingBox(None)
        if not bb:
            return start_point, start_point
        z_diff = bb.Max.Z - bb.Min.Z
        end_point = XYZ(start_point.X, start_point.Y, start_point.Z + z_diff)

    return start_point, end_point


def get_levels_user_choice():
    levels = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
    if not levels:
        forms.alert(u"В проекте нет уровней.", title=u"Ошибка", exitscript=True)

    levels_sorted = sorted(levels, key=lambda lvl: lvl.Elevation)
    items = []
    for lvl in levels_sorted:
        nm = lvl.Name
        idx = nm.lower().find(u"этаж")
        if idx != -1:
            nm = nm[:idx + len(u"этаж")]
        items.append(u"{0}  ({1:+.3f} м)".format(nm, UnitUtils.ConvertFromInternalUnits(lvl.Elevation, UnitTypeId.Meters)))

    chosen = forms.SelectFromList.show(items, multiselect=True, button_name=u"Выбрать уровни")
    if not chosen:
        sys.exit()

    result = []
    for lvl in levels_sorted:
        nm = lvl.Name
        idx = nm.lower().find(u"этаж")
        if idx != -1:
            nm_cut = nm[:idx + len(u"этаж")]
        else:
            nm_cut = nm
        item = u"{0}  ({1:+.3f} м)".format(nm_cut, UnitUtils.ConvertFromInternalUnits(lvl.Elevation, UnitTypeId.Meters))
        if item in chosen:
            result.append((nm_cut, lvl.Elevation, lvl))
    return result


def to_m_str(internal_ft):
    m = UnitUtils.ConvertFromInternalUnits(internal_ft, UnitTypeId.Meters)
    if abs(m) < 0.0005:
        return u"±0.000"
    return u"{:+.3f}".format(m)


def floor_title_from_level_name(level_name):
    """Возвращает строку вида "N этаж", где N — число, извлеченное из имени уровня.
    Примеры:
      'ВКК_01_Этаж' -> '1 этаж'
      'ВКК_11_Этаж' -> '11 этаж'
      'Этаж_05_+14.444' -> '5 этаж'
    Если число не найдено — возвращает исходное имя уровня.
    """
    try:
        s = level_name or u""
        # 1) Число ПЕРЕД словом 'этаж' (например '01_Этаж')
        m = re.search(u'(\d{1,3})\s*[_\-\s]*этаж', s, re.IGNORECASE | re.UNICODE)
        # 2) Число ПОСЛЕ слова 'этаж' (например 'Этаж_05')
        if not m:
            m = re.search(u'этаж\s*[_\-\s]*(\d{1,3})', s, re.IGNORECASE | re.UNICODE)
        # 3) Иначе — первое число из имени (избегая составных цифр), как фоллбэк
        if not m:
            m = re.search(u'(?<!\d)(\d{1,3})(?!\d)', s, re.UNICODE)
        if m:
            num = m.group(1)
            try:
                n = int(num.lstrip('0') or '0')
            except:
                n = int(num)
            return u"{0} этаж".format(n)
        return s
    except:
        return level_name


def main():
    start_up_checks()
    template = find_template_on_view()
    template_id = template.Id
    template_pt = template.Location.Point

    elements = get_pre_selected() or get_selected()
    tr = doc.ActiveProjectLocation.GetTransform()
    levels = get_levels_user_choice()

    # Строим кэш смещений для DisplacementElement
    build_displacement_cache()
    
    # Для каждого элемента определяем его смещение
    for el in elements:
        disp = find_element_displacement(el)
        if disp.GetLength() > TOL:
            displacement_cache[el.Id.IntegerValue] = disp

    placed = 0
    skipped = 0

    with revit.Transaction(u"Отметки уровней (без DLL)"):
        for el in elements:
            p1, p2 = get_connector_coordinates(el)
            if not p1 or not p2:
                skipped += 1
                continue

            # Получаем смещение для элемента (если он в наборе смещения)
            displacement = get_element_displacement(el.Id)

            z1_sh = tr.OfPoint(p1).Z
            z2_sh = tr.OfPoint(p2).Z
            zmin_sh = min(z1_sh, z2_sh)
            zmax_sh = max(z1_sh, z2_sh)

            for (lvl_name, lvl_elev_internal, lvl_obj) in levels:
                z_lvl_sh = tr.OfPoint(XYZ(0, 0, lvl_elev_internal)).Z
                if zmin_sh <= z_lvl_sh <= zmax_sh:
                    # Применяем смещение к позиции аннотации
                    new_pt = XYZ(p1.X + displacement.X, p1.Y + displacement.Y, lvl_elev_internal + displacement.Z)
                    move_vec = new_pt - template_pt
                    new_id = ElementTransformUtils.CopyElement(doc, template_id, move_vec)[0]
                    tag = doc.GetElement(new_id)
                    p_name = tag.LookupParameter(u"Имя уровня")
                    if p_name and not p_name.IsReadOnly:
                        # Вместо исходного имени записываем распарсенный вариант "N этаж"
                        parsed = floor_title_from_level_name(lvl_obj.Name)
                        p_name.Set(parsed)
                    p_elev = tag.LookupParameter(u"Отметка уровня")
                    if p_elev and not p_elev.IsReadOnly:
                        p_elev.Set(to_m_str(lvl_elev_internal))
                    placed += 1
                else:
                    skipped += 1

    # Подсчитываем смещённые элементы для информации
    displaced_count = sum(1 for el in elements if el.Id.IntegerValue in displacement_cache)
    
    msg = u"Готово. Размещено меток: {0}\nПропущено: {1}".format(placed, skipped)
    if displaced_count > 0:
        msg += u"\nЭлементов в наборах смещения: {0}".format(displaced_count)
    
    forms.alert(msg, title=u"Отметки уровней")

if __name__ == "__main__":
    main()
