# -*- coding: utf-8 -*-

import clr

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    View,
    ViewType,
    ElementId,
    ElementIdSetFilter,
    ParameterFilterElement,
    SelectionFilterElement,
    OverrideGraphicSettings,
    TemporaryViewMode,
    RevitLinkInstance,
    BuiltInCategory,
    View3D,
    ViewPlan,
    PlanViewRange,
    PlanViewPlane,
    XYZ,
    WorksetTable,
    WorksetKind,
    WorksetVisibility,
    LinkElementId,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

from pyrevit import script, forms

from System.Collections.Generic import List as Clist

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
output = script.get_output()

# Режим отладки - включить для диагностики
DEBUG_MODE = False

def debug_print(msg):
    """Вывод отладочной информации"""
    if DEBUG_MODE:
        print(msg)


def get_target_element():
    """
    Возвращает (element, linked_info):
      * element        - элемент в активном документе (в т.ч. экземпляр связи)
      * linked_info    - dict с данными о выбранном элементе из связи или None
    
    Если элемент из связи нужно проверить - пользователь должен выбрать его 
    заранее (Tab для переключения на элемент связи) перед запуском плагина.
    """
    sel = uidoc.Selection
    sel_ids = sel.GetElementIds()
    linked_info = None
    
    debug_print("=== ОТЛАДКА get_target_element ===")
    
    # Получаем количество выбранных элементов из активного документа
    count = 0
    try:
        count = sel_ids.Count
        debug_print("sel_ids.Count = {}".format(count))
    except Exception as e:
        debug_print("Ошибка получения Count: {}".format(e))
        count = 0

    # === ПРОВЕРКА ВЫБРАННЫХ ЭЛЕМЕНТОВ ИЗ СВЯЗЕЙ ===
    debug_print("Проверяем элементы из связей...")
    
    # Пробуем GetReferences - может содержать информацию о выбранных элементах из связей
    debug_print("Пробуем GetReferences()...")
    try:
        refs = sel.GetReferences()
        debug_print("  GetReferences вернул: {}".format(refs))
        if refs:
            debug_print("  Тип: {}".format(type(refs)))
            try:
                debug_print("  Count: {}".format(refs.Count))
            except:
                try:
                    debug_print("  len: {}".format(len(list(refs))))
                except:
                    pass
            
            for ref in refs:
                debug_print("  Reference:")
                debug_print("    ElementId: {}".format(ref.ElementId.IntegerValue if ref.ElementId else "None"))
                if hasattr(ref, 'LinkedElementId'):
                    debug_print("    LinkedElementId: {}".format(ref.LinkedElementId.IntegerValue if ref.LinkedElementId else "None"))
                    if ref.LinkedElementId and ref.LinkedElementId != ElementId.InvalidElementId:
                        # Нашли выбранный элемент из связи!
                        link_instance = doc.GetElement(ref.ElementId)
                        if isinstance(link_instance, RevitLinkInstance):
                            linked_elem_id = ref.LinkedElementId
                            link_doc = None
                            linked_elem = None
                            try:
                                link_doc = link_instance.GetLinkDocument()
                                if link_doc:
                                    linked_elem = link_doc.GetElement(linked_elem_id)
                            except:
                                pass
                            
                            linked_info = {
                                "link_instance": link_instance,
                                "linked_element_id": linked_elem_id,
                                "linked_element": linked_elem,
                                "link_doc": link_doc,
                            }
                            debug_print("  Найден элемент из связи через GetReferences!")
                            return link_instance, linked_info
    except Exception as e:
        debug_print("  Ошибка GetReferences: {}".format(e))
    
    # Пробуем другой подход - проверяем все связи через GetSelectedLinkedElements (если есть)
    debug_print("hasattr(sel, 'GetSelectedLinkedElements'): {}".format(hasattr(sel, 'GetSelectedLinkedElements')))
    
    if hasattr(sel, 'GetSelectedLinkedElements'):
        try:
            all_links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
            debug_print("Найдено связей в документе: {}".format(len(list(all_links))))
            
            for link in all_links:
                link_id = link.Id
                debug_print("  Проверяем связь ID: {}".format(link_id.IntegerValue))
                
                try:
                    linked_elem_ids = sel.GetSelectedLinkedElements(link_id)
                    debug_print("    GetSelectedLinkedElements вернул: {}".format(linked_elem_ids))
                    if linked_elem_ids and linked_elem_ids.Count > 0:
                        debug_print("    Найдены выбранные элементы в связи! Count: {}".format(linked_elem_ids.Count))
                        
                        # Получаем первый выбранный элемент
                        linked_elem_id = list(linked_elem_ids)[0]
                        debug_print("    linked_elem_id: {}".format(linked_elem_id.IntegerValue))
                        
                        link_doc = None
                        linked_elem = None
                        try:
                            link_doc = link.GetLinkDocument()
                            if link_doc:
                                linked_elem = link_doc.GetElement(linked_elem_id)
                                debug_print("    linked_elem: {}".format(type(linked_elem).__name__ if linked_elem else "None"))
                                if linked_elem and linked_elem.Category:
                                    debug_print("    Category: {}".format(linked_elem.Category.Name))
                        except Exception as e:
                            debug_print("    Ошибка получения элемента: {}".format(e))
                        
                        linked_info = {
                            "link_instance": link,
                            "linked_element_id": linked_elem_id,
                            "linked_element": linked_elem,
                            "link_doc": link_doc,
                        }
                        debug_print("  Возвращаем link + linked_info")
                        return link, linked_info
                except Exception as e:
                    debug_print("    Ошибка GetSelectedLinkedElements: {}".format(e))
        except Exception as e:
            debug_print("Ошибка при проверке связей: {}".format(e))

    # === ПРОВЕРКА ОБЫЧНЫХ ЭЛЕМЕНТОВ ===
    if count > 0:
        debug_print("Есть выбранные элементы в активном документе, обрабатываем...")
        for elem_id in sel_ids:
            elem = doc.GetElement(elem_id)
            if elem is None:
                continue
            
            debug_print("  Элемент: {} (ID: {})".format(type(elem).__name__, elem_id.IntegerValue))
            
            # Проверяем, является ли это экземпляром связи
            if isinstance(elem, RevitLinkInstance):
                debug_print("  -> Это RevitLinkInstance (выбрана сама связь)")
                link_doc = None
                try:
                    link_doc = elem.GetLinkDocument()
                except:
                    link_doc = None
                
                linked_info = {
                    "link_instance": elem,
                    "linked_element_id": None,
                    "linked_element": None,
                    "link_doc": link_doc,
                }
                return elem, linked_info
            else:
                debug_print("  -> Обычный элемент, возвращаем")
                return elem, None
    
    # === ЕСЛИ НИЧЕГО НЕ ВЫБРАНО ===
    debug_print("Ничего не выбрано, вызываем PickObject...")
    try:
        ref = sel.PickObject(
            ObjectType.Element,
            "Выберите элемент для проверки фильтров на активном виде"
        )
        if ref:
            elem = doc.GetElement(ref.ElementId)
            if elem is not None:
                return elem, None
    except Exception as e:
        debug_print("Ошибка PickObject: {}".format(e))
        return None, None

    debug_print("Возвращаем None, None")
    return None, None


def has_any_override(ogs):
    """Пытается определить, есть ли какие-либо графические переопределения в OverrideGraphicSettings."""
    if ogs is None:
        return False

    # Цвет линий
    try:
        col = getattr(ogs, "ProjectionLineColor", None)
        if col is not None:
            try:
                if col.IsValid:
                    return True
            except:
                try:
                    if col.Red != 0 or col.Green != 0 or col.Blue != 0:
                        return True
                except:
                    pass
    except:
        pass

    element_id_type = ElementId

    # Штриховки / паттерны
    ids_to_check = [
        getattr(ogs, "ProjectionLinePatternId", None),
        getattr(ogs, "CutLinePatternId", None),
        getattr(ogs, "SurfaceForegroundPatternId", None),
        getattr(ogs, "SurfaceBackgroundPatternId", None),
        getattr(ogs, "CutForegroundPatternId", None),
        getattr(ogs, "CutBackgroundPatternId", None),
    ]

    for pid in ids_to_check:
        try:
            if isinstance(pid, element_id_type):
                if pid.IntegerValue != -1:
                    return True
        except:
            pass

    # Толщина линий
    for wprop in ("ProjectionLineWeight", "CutLineWeight"):
        try:
            w = getattr(ogs, wprop, None)
            if w is not None and w > 0:
                return True
        except:
            pass

    return False


def _get_element_category_for_filters(element):
    """
    Возвращает категорию элемента для анализа фильтров.

    Для экземпляров связи пытаемся явно взять категорию 'Revit Links', если element.Category == None.
    """
    if element is None:
        return None

    category = element.Category
    if category is not None:
        return category

    # Специальный кейс для RevitLinkInstance
    try:
        if isinstance(element, RevitLinkInstance):
            try:
                cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_RvtLinks)
                if cat is not None:
                    return cat
            except:
                return None
    except:
        pass

    return None


def find_filters_for_element(view, element, linked_info=None):
    """Находит фильтры вида, которые реально отбирают данный элемент."""
    result = []

    if view is None or element is None:
        return result

    # Для элементов из связи используем данные из linked_info
    is_linked = linked_info and linked_info.get("linked_element")
    
    if is_linked:
        linked_element = linked_info.get("linked_element")
        linked_element_id = linked_info.get("linked_element_id")
        link_doc = linked_info.get("link_doc")
        category = _get_element_category_for_filters(linked_element)
        
        # Для элементов из связи создаём фильтр по ID элемента в документе связи
        linked_id_list = Clist[ElementId]()
        linked_id_list.Add(linked_element_id)
        linked_elem_id_filter = ElementIdSetFilter(linked_id_list)
    else:
        element_id = element.Id
        category = _get_element_category_for_filters(element)
        
        # Подготавливаем фильтр по конкретному элементу
        id_list = Clist[ElementId]()
        id_list.Add(element_id)
        elem_id_filter = ElementIdSetFilter(id_list)

    if category is None:
        return result

    try:
        filter_ids = list(view.GetFilters())
    except:
        filter_ids = []

    if not filter_ids:
        return result

    for fid in filter_ids:
        try:
            f_elem = doc.GetElement(fid)
        except:
            continue

        if f_elem is None:
            continue

        filter_type = f_elem.GetType().Name

        # Общая видимость фильтра на виде
        try:
            visible_by_filter = view.GetFilterVisibility(fid)
        except:
            visible_by_filter = True

        # Фильтр по параметрам
        if isinstance(f_elem, ParameterFilterElement):
            try:
                cats = f_elem.GetCategories()
            except:
                cats = None

            if cats and not cats.Contains(category.Id):
                # Категория элемента не входит в категории фильтра
                continue

            passes_filter = False
            try:
                elem_filter = f_elem.GetElementFilter()
            except:
                elem_filter = None

            if elem_filter is not None:
                # Проверяем, отбирают ли правила фильтра именно этот элемент
                # ВАЖНО: используем коллектор по документу, а не по виду,
                # т.к. элемент может быть уже скрыт фильтром на виде
                try:
                    if is_linked and link_doc:
                        # Для элемента из связи проверяем в документе связи
                        collector = FilteredElementCollector(link_doc)
                        collector = collector.WherePasses(elem_filter).WherePasses(linked_elem_id_filter)
                    else:
                        # Для обычного элемента проверяем в основном документе
                        collector = FilteredElementCollector(doc)
                        collector = collector.WherePasses(elem_filter).WherePasses(elem_id_filter)
                    if collector.GetElementCount() > 0:
                        passes_filter = True
                except:
                    passes_filter = False
            else:
                # Нет ElementFilter'а (редкий случай) - считаем, что фильтр влияет по категории
                passes_filter = True

            if not passes_filter:
                continue

            # Проверяем, есть ли графические переопределения у фильтра на виде
            try:
                ogs = view.GetFilterOverrides(fid)
            except:
                ogs = None

            has_override = False
            try:
                has_override = has_any_override(ogs)
            except:
                has_override = False

            # Список категорий по имени
            cat_names = []
            if cats:
                for cid in cats:
                    try:
                        cat = doc.Settings.Categories.get_Item(cid)
                    except:
                        cat = None
                    if cat:
                        try:
                            cat_names.append(cat.Name)
                        except:
                            continue

            result.append({
                "filter": f_elem,
                "type": filter_type,
                "categories": cat_names,
                "has_override": has_override,
                "visible": visible_by_filter,
            })

        # Фильтр по выбору
        elif isinstance(f_elem, SelectionFilterElement):
            passes_filter = False
            try:
                if hasattr(f_elem, "AllowsElement"):
                    if f_elem.AllowsElement(element):
                        passes_filter = True
                else:
                    elem_ids = f_elem.GetElementIds()
                    if elem_ids and element_id in elem_ids:
                        passes_filter = True
            except:
                passes_filter = False

            if not passes_filter:
                continue

            # Для SelectionFilterElement графические настройки обычно не используются,
            # но на всякий случай проверим GetFilterOverrides.
            try:
                ogs = view.GetFilterOverrides(fid)
            except:
                ogs = None

            has_override = False
            try:
                has_override = has_any_override(ogs)
            except:
                has_override = False

            result.append({
                "filter": f_elem,
                "type": filter_type,
                "categories": [],
                "has_override": has_override,
                "visible": visible_by_filter,
            })

        else:
            # Другие типы фильтров - пропускаем
            continue

    return result


def _get_world_box_from_view_box(box):
    """
    Строит мировой AABB для CropBox / SectionBox:
    берём 8 углов локального бокса, умножаем на Transform и строим общий min/max.
    """
    if box is None:
        return None

    try:
        bb_min = box.Min
        bb_max = box.Max
        t = box.Transform
    except:
        return None

    pts = [
        XYZ(bb_min.X, bb_min.Y, bb_min.Z),
        XYZ(bb_min.X, bb_min.Y, bb_max.Z),
        XYZ(bb_min.X, bb_max.Y, bb_min.Z),
        XYZ(bb_min.X, bb_max.Y, bb_max.Z),
        XYZ(bb_max.X, bb_min.Y, bb_min.Z),
        XYZ(bb_max.X, bb_min.Y, bb_max.Z),
        XYZ(bb_max.X, bb_max.Y, bb_min.Z),
        XYZ(bb_max.X, bb_max.Y, bb_max.Z),
    ]

    xs = []
    ys = []
    zs = []
    for p in pts:
        try:
            wp = t.OfPoint(p)
        except:
            continue
        xs.append(wp.X)
        ys.append(wp.Y)
        zs.append(wp.Z)

    if not xs or not ys or not zs:
        return None

    class MinMax(object):
        pass

    world_min = MinMax()
    world_min.X = min(xs)
    world_min.Y = min(ys)
    world_min.Z = min(zs)

    world_max = MinMax()
    world_max.X = max(xs)
    world_max.Y = max(ys)
    world_max.Z = max(zs)

    class WBox(object):
        def __init__(self, mn, mx):
            self.Min = mn
            self.Max = mx

    return WBox(world_min, world_max)


def _aabb_intersects(a_min, a_max, b_min, b_max):
    """Проверка пересечения двух мировых AABB."""
    try:
        no_intersection = (
            a_max.X < b_min.X or a_min.X > b_max.X or
            a_max.Y < b_min.Y or a_min.Y > b_max.Y or
            a_max.Z < b_min.Z or a_min.Z > b_max.Z
        )
    except:
        return None
    return not no_intersection


def _aabb_intersects_2d(a_min, a_max, b_min, b_max):
    """Проверка пересечения двух AABB только по X и Y (для планов)."""
    try:
        no_intersection = (
            a_max.X < b_min.X or a_min.X > b_max.X or
            a_max.Y < b_min.Y or a_min.Y > b_max.Y
        )
    except:
        return None
    return not no_intersection


def _get_plan_view_range_z(view):
    """
    Получает границы секущего диапазона (View Range) для плана в мировых координатах Z.
    Возвращает (z_bottom, z_top) или None.
    """
    try:
        if not isinstance(view, ViewPlan):
            return None
        
        vr = view.GetViewRange()
        if vr is None:
            return None
        
        # Получаем уровень вида
        level = view.GenLevel
        if level is None:
            return None
        level_elevation = level.Elevation
        
        # Получаем смещения относительно уровня
        # Bottom - нижняя граница (обычно View Depth или Bottom)
        # Top - верхняя граница (обычно Cut Plane или Top)
        
        # Для корректной работы используем Bottom и Top Clip Plane
        bottom_offset = vr.GetOffset(PlanViewPlane.BottomClipPlane)
        top_offset = vr.GetOffset(PlanViewPlane.TopClipPlane)
        
        # Проверяем уровни, к которым привязаны плоскости
        bottom_level_id = vr.GetLevelId(PlanViewPlane.BottomClipPlane)
        top_level_id = vr.GetLevelId(PlanViewPlane.TopClipPlane)
        
        # Вычисляем абсолютные Z-координаты
        if bottom_level_id and bottom_level_id.IntegerValue != -1:
            bottom_level = doc.GetElement(bottom_level_id)
            if bottom_level:
                z_bottom = bottom_level.Elevation + bottom_offset
            else:
                z_bottom = level_elevation + bottom_offset
        else:
            z_bottom = level_elevation + bottom_offset
        
        if top_level_id and top_level_id.IntegerValue != -1:
            top_level = doc.GetElement(top_level_id)
            if top_level:
                z_top = top_level.Elevation + top_offset
            else:
                z_top = level_elevation + top_offset
        else:
            z_top = level_elevation + top_offset
        
        # Также учитываем View Depth (глубина вида) - самая нижняя видимая точка
        try:
            view_depth_offset = vr.GetOffset(PlanViewPlane.ViewDepthPlane)
            view_depth_level_id = vr.GetLevelId(PlanViewPlane.ViewDepthPlane)
            
            if view_depth_level_id and view_depth_level_id.IntegerValue != -1:
                view_depth_level = doc.GetElement(view_depth_level_id)
                if view_depth_level:
                    z_view_depth = view_depth_level.Elevation + view_depth_offset
                else:
                    z_view_depth = level_elevation + view_depth_offset
            else:
                z_view_depth = level_elevation + view_depth_offset
            
            # View Depth может быть ниже Bottom Clip Plane
            z_bottom = min(z_bottom, z_view_depth)
        except:
            pass
        
        return (z_bottom, z_top)
    except:
        return None


def is_inside_crop(view, element):
    """
    Пытается определить, попадает ли элемент в объём подрезки вида / секущий диапазон / 3D-бокс.

    Логика для разных типов видов:
    - Для планов: проверяем CropBox по X,Y и ViewRange по Z
    - Для 3D: проверяем SectionBox (если активен) или CropBox
    - Для разрезов/фасадов: проверяем CropBox в мировых координатах
    """
    if view is None or element is None:
        return None

    # Получаем bounding box элемента в мировых координатах
    elem_bb = None
    try:
        elem_bb = element.get_BoundingBox(None)
    except:
        elem_bb = None

    if elem_bb is None:
        try:
            elem_bb = element.get_BoundingBox(view)
        except:
            elem_bb = None

    if elem_bb is None:
        return None

    try:
        eb_min = elem_bb.Min
        eb_max = elem_bb.Max
    except:
        return None

    # === Обработка планов (ViewPlan) ===
    try:
        if isinstance(view, ViewPlan):
            # 1. Проверяем CropBox по X и Y (если активен)
            crop_ok = True
            try:
                if hasattr(view, "CropBoxActive") and view.CropBoxActive:
                    crop_box = view.CropBox
                    if crop_box is not None:
                        # Для плана CropBox уже в мировых координатах (или близко к ним)
                        # Transform обычно единичный или поворот вокруг Z
                        world_crop = _get_world_box_from_view_box(crop_box)
                        if world_crop is not None:
                            crop_ok = _aabb_intersects_2d(eb_min, eb_max, world_crop.Min, world_crop.Max)
                            if crop_ok is False:
                                return False
            except:
                pass
            
            # 2. Проверяем ViewRange по Z
            z_range = _get_plan_view_range_z(view)
            if z_range is not None:
                z_bottom, z_top = z_range
                # Элемент должен пересекаться с диапазоном по Z
                if eb_max.Z < z_bottom or eb_min.Z > z_top:
                    return False
            
            return True if crop_ok else None
    except:
        pass

    # === Обработка 3D-видов ===
    try:
        if isinstance(view, View3D):
            view_box = None
            
            # Сначала проверяем SectionBox
            try:
                is_section_active = False
                if hasattr(view, "IsSectionBoxActive"):
                    is_section_active = view.IsSectionBoxActive
                
                if is_section_active:
                    view_box = view.GetSectionBox()
                    if view_box is None:
                        view_box = view.SectionBox
            except:
                pass
            
            # Если SectionBox не активен, проверяем CropBox
            if view_box is None:
                try:
                    if hasattr(view, "CropBoxActive") and view.CropBoxActive:
                        view_box = view.CropBox
                except:
                    pass
            
            if view_box is None:
                # Нет ни SectionBox, ни CropBox - элемент виден (по этому критерию)
                return None
            
            # Преобразуем в мировые координаты и проверяем пересечение
            world_view_box = _get_world_box_from_view_box(view_box)
            if world_view_box is None:
                return None
            
            return _aabb_intersects(eb_min, eb_max, world_view_box.Min, world_view_box.Max)
    except:
        pass

    # === Обработка разрезов, фасадов и других видов ===
    view_box = None
    try:
        if hasattr(view, "CropBoxActive") and not view.CropBoxActive:
            return None  # CropBox не активен
    except:
        pass
    
    try:
        view_box = view.CropBox
    except:
        view_box = None

    if view_box is None:
        return None

    # Преобразуем бокс вида в мировой AABB
    world_view_box = _get_world_box_from_view_box(view_box)
    if world_view_box is None:
        return None

    return _aabb_intersects(eb_min, eb_max, world_view_box.Min, world_view_box.Max)


def is_inside_crop_linked(view, linked_element, link_instance):
    """
    Проверяет, попадает ли элемент из связи в объём подрезки вида.
    Учитывает трансформацию экземпляра связи.
    """
    if view is None or linked_element is None or link_instance is None:
        return None
    
    # Получаем bounding box элемента в координатах связанного документа
    elem_bb = None
    try:
        elem_bb = linked_element.get_BoundingBox(None)
    except:
        pass
    
    if elem_bb is None:
        return None
    
    # Получаем трансформацию экземпляра связи
    try:
        link_transform = link_instance.GetTotalTransform()
    except:
        link_transform = None
    
    # Трансформируем bounding box элемента в мировые координаты
    try:
        if link_transform:
            # Трансформируем Min и Max точки
            eb_min_local = elem_bb.Min
            eb_max_local = elem_bb.Max
            
            # Получаем все 8 углов и трансформируем их
            corners = [
                XYZ(eb_min_local.X, eb_min_local.Y, eb_min_local.Z),
                XYZ(eb_min_local.X, eb_min_local.Y, eb_max_local.Z),
                XYZ(eb_min_local.X, eb_max_local.Y, eb_min_local.Z),
                XYZ(eb_min_local.X, eb_max_local.Y, eb_max_local.Z),
                XYZ(eb_max_local.X, eb_min_local.Y, eb_min_local.Z),
                XYZ(eb_max_local.X, eb_min_local.Y, eb_max_local.Z),
                XYZ(eb_max_local.X, eb_max_local.Y, eb_min_local.Z),
                XYZ(eb_max_local.X, eb_max_local.Y, eb_max_local.Z),
            ]
            
            xs, ys, zs = [], [], []
            for corner in corners:
                transformed = link_transform.OfPoint(corner)
                xs.append(transformed.X)
                ys.append(transformed.Y)
                zs.append(transformed.Z)
            
            class MinMax:
                pass
            
            eb_min = MinMax()
            eb_min.X = min(xs)
            eb_min.Y = min(ys)
            eb_min.Z = min(zs)
            
            eb_max = MinMax()
            eb_max.X = max(xs)
            eb_max.Y = max(ys)
            eb_max.Z = max(zs)
        else:
            eb_min = elem_bb.Min
            eb_max = elem_bb.Max
    except:
        eb_min = elem_bb.Min
        eb_max = elem_bb.Max
    
    # Далее используем ту же логику, что и в is_inside_crop
    # === Обработка планов (ViewPlan) ===
    try:
        if isinstance(view, ViewPlan):
            crop_ok = True
            try:
                if hasattr(view, "CropBoxActive") and view.CropBoxActive:
                    crop_box = view.CropBox
                    if crop_box is not None:
                        world_crop = _get_world_box_from_view_box(crop_box)
                        if world_crop is not None:
                            crop_ok = _aabb_intersects_2d(eb_min, eb_max, world_crop.Min, world_crop.Max)
                            if crop_ok is False:
                                return False
            except:
                pass
            
            z_range = _get_plan_view_range_z(view)
            if z_range is not None:
                z_bottom, z_top = z_range
                if eb_max.Z < z_bottom or eb_min.Z > z_top:
                    return False
            
            return True if crop_ok else None
    except:
        pass

    # === Обработка 3D-видов ===
    try:
        if isinstance(view, View3D):
            view_box = None
            
            try:
                is_section_active = False
                if hasattr(view, "IsSectionBoxActive"):
                    is_section_active = view.IsSectionBoxActive
                
                if is_section_active:
                    view_box = view.GetSectionBox()
                    if view_box is None:
                        view_box = view.SectionBox
            except:
                pass
            
            if view_box is None:
                try:
                    if hasattr(view, "CropBoxActive") and view.CropBoxActive:
                        view_box = view.CropBox
                except:
                    pass
            
            if view_box is None:
                return None
            
            world_view_box = _get_world_box_from_view_box(view_box)
            if world_view_box is None:
                return None
            
            return _aabb_intersects(eb_min, eb_max, world_view_box.Min, world_view_box.Max)
    except:
        pass

    # === Обработка разрезов, фасадов и других видов ===
    view_box = None
    try:
        if hasattr(view, "CropBoxActive") and not view.CropBoxActive:
            return None
    except:
        pass
    
    try:
        view_box = view.CropBox
    except:
        view_box = None

    if view_box is None:
        return None

    world_view_box = _get_world_box_from_view_box(view_box)
    if world_view_box is None:
        return None

    return _aabb_intersects(eb_min, eb_max, world_view_box.Min, world_view_box.Max)


def get_visibility_info(view, element, linked_info=None):
    """Собирает дополнительную информацию о видимости и переопределениях для элемента и его категории."""
    info = {}

    if view is None or element is None:
        return info

    # Определяем, с каким элементом работаем для проверки категории
    # Для элементов из связи - используем linked_element
    linked_element = None
    link_instance = None
    if linked_info:
        linked_element = linked_info.get("linked_element")
        link_instance = linked_info.get("link_instance")

    # Категория для проверки - либо элемента из связи, либо обычного элемента
    check_element = linked_element if linked_element else element
    
    category = _get_element_category_for_filters(check_element)
    if category is not None:
        try:
            info["category_name"] = category.Name
        except:
            info["category_name"] = None

        # Скрыта ли категория на виде
        try:
            hidden_cat = view.GetCategoryHidden(category.Id)
            info["category_hidden"] = bool(hidden_cat)
        except:
            info["category_hidden"] = None

        # Переопределения категории на виде
        try:
            cat_ogs = view.GetCategoryOverrides(category.Id)
        except:
            cat_ogs = None

        has_cat_overrides = False
        try:
            has_cat_overrides = has_any_override(cat_ogs)
        except:
            has_cat_overrides = False

        info["category_has_overrides"] = has_cat_overrides
    else:
        info["category_name"] = None
        info["category_hidden"] = None
        info["category_has_overrides"] = False

    # Постоянно скрыт ли элемент на виде (Hide in View / Скрыть при просмотре)
    # Для связей проверяем сам экземпляр связи
    try:
        info["element_hidden"] = element.IsHidden(view)
    except:
        info["element_hidden"] = None

    # Переопределения конкретного элемента (экземпляра связи)
    try:
        elem_ogs = view.GetElementOverrides(element.Id)
    except:
        elem_ogs = None

    has_elem_overrides = False
    try:
        has_elem_overrides = has_any_override(elem_ogs)
    except:
        has_elem_overrides = False

    info["element_has_overrides"] = has_elem_overrides

    # Временный режим вида (изоляция/скрытие, Показать скрытые элементы и т.п.)
    try:
        tvm = view.TemporaryViewMode
        info["temporary_view_mode"] = str(tvm)
    except:
        info["temporary_view_mode"] = None

    # Попадает ли элемент в объём подрезки вида / секущий диапазон / 3D-бокс
    # Для элементов из связи - проверяем с учётом трансформации связи
    try:
        if linked_element and link_instance:
            inside_crop = is_inside_crop_linked(view, linked_element, link_instance)
        else:
            inside_crop = is_inside_crop(view, element)
    except:
        inside_crop = None

    info["inside_crop"] = inside_crop

    # Дополнительно для связей - проверяем видимость категории "Связанные файлы"
    if linked_info:
        link_category = _get_element_category_for_filters(element)
        if link_category:
            try:
                info["link_category_name"] = link_category.Name
                info["link_category_hidden"] = bool(view.GetCategoryHidden(link_category.Id))
            except:
                pass
        
        # Проверяем, скрыт ли сам экземпляр связи на виде
        try:
            info["link_instance_hidden"] = element.IsHidden(view)
        except:
            info["link_instance_hidden"] = None
        
        # Примечание: для элементов из связи нет надёжного API для проверки IsHidden.
        # Определяем скрытие элемента из связи косвенно:
        # если включён режим "Показать скрытые элементы" и других причин скрытия нет,
        # значит элемент скрыт через "Скрыть при просмотре".
        # Это делается в show_result().

    # === Проверка рабочих наборов ===
    # Проверяем, включена ли работа с рабочими наборами в документе
    try:
        if doc.IsWorkshared:
            workset_table = doc.GetWorksetTable()
            
            # Для элемента из связи проверяем рабочий набор linked_element
            check_elem_for_workset = linked_element if linked_element else element
            check_doc_for_workset = linked_info.get("link_doc") if linked_info and linked_info.get("link_doc") else doc
            
            # Получаем рабочий набор элемента
            try:
                workset_id = check_elem_for_workset.WorksetId
                if workset_id and workset_id.IntegerValue != -1:
                    # Получаем информацию о рабочем наборе
                    if check_doc_for_workset and check_doc_for_workset.IsWorkshared:
                        ws_table = check_doc_for_workset.GetWorksetTable()
                        workset = ws_table.GetWorkset(workset_id)
                        if workset:
                            info["workset_name"] = workset.Name
                            info["workset_is_open"] = workset.IsOpen
                            
                            # Проверяем видимость рабочего набора на виде
                            try:
                                ws_visibility = view.GetWorksetVisibility(workset_id)
                                if ws_visibility == WorksetVisibility.Hidden:
                                    info["workset_visible_on_view"] = False
                                elif ws_visibility == WorksetVisibility.Visible:
                                    info["workset_visible_on_view"] = True
                                else:
                                    # UseGlobalSetting - используем глобальную настройку
                                    info["workset_visible_on_view"] = workset.IsVisibleByDefault
                            except:
                                info["workset_visible_on_view"] = None
            except:
                pass
            
            # Для связей - также проверяем рабочий набор самой связи
            if linked_info and element:
                try:
                    link_workset_id = element.WorksetId
                    if link_workset_id and link_workset_id.IntegerValue != -1:
                        link_workset = workset_table.GetWorkset(link_workset_id)
                        if link_workset:
                            info["link_workset_name"] = link_workset.Name
                            info["link_workset_is_open"] = link_workset.IsOpen
                            
                            # Проверяем видимость рабочего набора связи на виде
                            try:
                                link_ws_visibility = view.GetWorksetVisibility(link_workset_id)
                                if link_ws_visibility == WorksetVisibility.Hidden:
                                    info["link_workset_visible_on_view"] = False
                                elif link_ws_visibility == WorksetVisibility.Visible:
                                    info["link_workset_visible_on_view"] = True
                                else:
                                    info["link_workset_visible_on_view"] = link_workset.IsVisibleByDefault
                            except:
                                info["link_workset_visible_on_view"] = None
                except:
                    pass
    except:
        pass

    return info


def show_result(filters_info, visibility_info, view, element, linked_info=None):
    """Выводит результаты в окно вывода pyRevit в виде Markdown-таблиц."""
    output.print_md(u"---")
    output.print_md(u"## 🔍 К каким фильтрам относится элемент?")
    output.print_md(u"")

    try:
        view_title = u"{} ({})".format(view.Name, view.ViewType)
    except:
        view_title = view.Name

    try:
        elem_id_str = str(element.Id.IntegerValue)
    except:
        elem_id_str = str(element.Id)

    output.print_md(u"**Активный вид:** {}  ".format(view_title))
    output.print_md(u"**ID элемента (в активном файле):** `{}`  ".format(elem_id_str))

    # Доп. информация, если элемент выбран из связи
    if linked_info:
        link_instance = linked_info.get("link_instance")
        linked_element = linked_info.get("linked_element")
        linked_id = linked_info.get("linked_element_id")
        link_doc = linked_info.get("link_doc")

        output.print_md(u"")
        output.print_md(u"### 🔗 Элемент из связанной модели")

        try:
            link_name = link_instance.Name if link_instance else u"—"
        except:
            link_name = u"—"

        try:
            link_doc_title = link_doc.Title if link_doc else u"—"
        except:
            link_doc_title = u"—"

        try:
            linked_id_str = str(linked_id.IntegerValue) if linked_id else u"—"
        except:
            linked_id_str = u"—"

        try:
            linked_cat = linked_element.Category.Name if linked_element and linked_element.Category else u"—"
        except:
            linked_cat = u"—"

        link_table_data = [
            [u"Связь", link_name],
            [u"Файл связи", link_doc_title],
            [u"ID элемента в связи", linked_id_str],
            [u"Категория элемента", linked_cat],
        ]
        
        # Показываем статус скрытия экземпляра связи
        link_instance_hidden = visibility_info.get("link_instance_hidden")
        if link_instance_hidden is True:
            link_table_data.append([u"⚠️ Экземпляр связи", u"СКРЫТ на виде"])
        elif link_instance_hidden is False:
            link_table_data.append([u"Экземпляр связи", u"отображается"])
        
        # Показываем статус видимости категории
        cat_hidden = visibility_info.get("category_hidden")
        if cat_hidden is True:
            link_table_data.append([u"⚠️ Видимость категории", u"СКРЫТА на виде"])
        elif cat_hidden is False:
            link_table_data.append([u"Видимость категории", u"отображается"])
        
        # Показываем статус категории "Связанные файлы"
        link_cat_hidden = visibility_info.get("link_category_hidden")
        link_cat_name = visibility_info.get("link_category_name", u"Связанные файлы")
        if link_cat_hidden is True:
            link_table_data.append([u"⚠️ Категория «{}»".format(link_cat_name), u"СКРЫТА на виде"])
        elif link_cat_hidden is False:
            link_table_data.append([u"Категория «{}»".format(link_cat_name), u"отображается"])
        
        # Показываем статус скрытия элемента из связи
        linked_elem_hidden = visibility_info.get("linked_element_hidden")
        if linked_elem_hidden is True:
            link_table_data.append([u"⚠️ Элемент в связи", u"СКРЫТ на виде (Скрыть при просмотре)"])
        elif linked_elem_hidden is False:
            link_table_data.append([u"Элемент в связи", u"отображается"])
        
        output.print_table(link_table_data, columns=[u"Параметр", u"Значение"])
        output.print_md(u"")

    output.print_md(u"")

    # --- Блок причин невидимости ---
    cat_hidden = visibility_info.get("category_hidden")
    cat_name = visibility_info.get("category_name")
    elem_hidden = visibility_info.get("element_hidden")
    link_cat_hidden = visibility_info.get("link_category_hidden")
    link_cat_name = visibility_info.get("link_category_name")
    link_instance_hidden = visibility_info.get("link_instance_hidden")
    linked_element_hidden = visibility_info.get("linked_element_hidden")
    tvm = visibility_info.get("temporary_view_mode")
    inside_crop = visibility_info.get("inside_crop")
    
    # Рабочие наборы
    workset_name = visibility_info.get("workset_name")
    workset_is_open = visibility_info.get("workset_is_open")
    workset_visible = visibility_info.get("workset_visible_on_view")
    link_workset_name = visibility_info.get("link_workset_name")
    link_workset_is_open = visibility_info.get("link_workset_is_open")
    link_workset_visible = visibility_info.get("link_workset_visible_on_view")
    
    reveal_mode = False
    if tvm and "RevealHiddenElements" in tvm:
        reveal_mode = True

    # Фильтры, которые выключают видимость (visible == False)
    hiding_filters = [f for f in filters_info if not f.get("visible", True)]

    reasons = []

    # Проверка скрытия экземпляра связи
    if linked_info and link_instance_hidden:
        if reveal_mode:
            reasons.append(u"🔗 <span style='color:#e74c3c;'>Экземпляр связи скрыт на виде (Скрыть при просмотре), сейчас виден только в режиме «Показать скрытые элементы».</span>")
        else:
            reasons.append(u"🔗 <span style='color:#e74c3c;'>Экземпляр связи скрыт на виде (Скрыть при просмотре).</span>")

    if elem_hidden:
        if reveal_mode:
            reasons.append(u"🔴 <span style='color:#e74c3c;'>Элемент скрыт на виде (Скрыть при просмотре), сейчас отображается только из-за режима «Показать скрытые элементы».</span>")
        else:
            reasons.append(u"🔴 <span style='color:#e74c3c;'>Элемент скрыт на виде (Скрыть при просмотре).</span>")

    if cat_hidden:
        cat_name_str = u"«{}»".format(cat_name) if cat_name else u"элемента"
        if reveal_mode:
            reasons.append(u"🟥 <span style='color:#e74c3c;'>Категория {} скрыта на виде, видна только в режиме «Показать скрытые элементы».</span>".format(cat_name_str))
        else:
            reasons.append(u"🟥 <span style='color:#e74c3c;'>Категория {} скрыта на виде.</span>".format(cat_name_str))

    # Проверка категории "Связанные файлы" для элементов из связи
    if link_cat_hidden:
        link_cat_str = u"«{}»".format(link_cat_name) if link_cat_name else u"«Связанные файлы»"
        if reveal_mode:
            reasons.append(u"🟥 <span style='color:#e74c3c;'>Категория {} скрыта на виде (связь не отображается), видна только в режиме «Показать скрытые элементы».</span>".format(link_cat_str))
        else:
            reasons.append(u"🟥 <span style='color:#e74c3c;'>Категория {} скрыта на виде (связь не отображается).</span>".format(link_cat_str))

    # Проверка рабочих наборов
    if workset_is_open is False:
        ws_name_str = u"«{}»".format(workset_name) if workset_name else u"элемента"
        reasons.append(u"📁 <span style='color:#e74c3c;'>Рабочий набор {} закрыт (не загружен в память).</span>".format(ws_name_str))
    
    if workset_visible is False:
        ws_name_str = u"«{}»".format(workset_name) if workset_name else u"элемента"
        reasons.append(u"📁 <span style='color:#e74c3c;'>Рабочий набор {} скрыт на виде (Видимость рабочих наборов).</span>".format(ws_name_str))

    # Проверка рабочего набора связи
    if link_workset_is_open is False:
        link_ws_name_str = u"«{}»".format(link_workset_name) if link_workset_name else u"связи"
        reasons.append(u"📁 <span style='color:#e74c3c;'>Рабочий набор {} (связь) закрыт (не загружен в память).</span>".format(link_ws_name_str))
    
    if link_workset_visible is False:
        link_ws_name_str = u"«{}»".format(link_workset_name) if link_workset_name else u"связи"
        reasons.append(u"📁 <span style='color:#e74c3c;'>Рабочий набор {} (связь) скрыт на виде (Видимость рабочих наборов).</span>".format(link_ws_name_str))

    if hiding_filters:
        names = []
        for finfo in hiding_filters:
            f = finfo["filter"]
            try:
                names.append(f.Name)
            except:
                names.append(u"<без имени>")
        names_str = u", ".join(names)
        reasons.append(u"🚫 <span style='color:#e74c3c;'>Видимость элемента отключена фильтром(ами) вида: {}</span>".format(names_str))

    # Объём подрезки / секущий диапазон / глубина / 3D-бокс
    if inside_crop is False:
        reasons.append(u"📦 <span style='color:#e74c3c;'>Элемент не попадает в объём подрезки вида (секущий диапазон / глубина / объём подрезки 3D-вида).</span>")

    if reveal_mode and not reasons:
        reasons.append(u"🟠 Режим «Показать скрытые элементы» включён, но явных причин скрытия элемента не обнаружено (категория и сам элемент не скрыты, фильтры не отключают видимость, по объёму подрезки элемент, вероятно, попадает в видимость).")

    if not reasons:
        reasons.append(u"🟢 <span style='color:#27ae60;'>Явных причин невидимости элемента на виде не обнаружено (категория и элемент не скрыты, фильтры не отключают видимость, элемент попадает в объём подрезки вида).</span>")
    
    # Дисклеймер для элементов из связи
    if linked_info and not hiding_filters and not cat_hidden and not link_cat_hidden and inside_crop is not False:
        reasons.append(u"🟡 <span style='color:#f39c12;'>Если элемент из связи всё ещё не отображается на виде, скорее всего он скрыт через «Скрыть при просмотре». Включите режим «Показать скрытые элементы» и выберите «Показать при просмотре» для этого элемента. Либо категория элемента выключена в Переопределении видимости ДЛЯ СВЯЗИ</span>")

    output.print_md(u"### ⚠️ Причины невидимости элемента")
    for r in reasons:
        output.print_md(u"- {}".format(r))
    output.print_md(u"")

    # --- Таблица фильтров ---
    if filters_info:
        output.print_md(u"### ✅ Фильтры вида, влияющие на элемент")
        output.print_md(u"")
        
        filters_table_data = []
        for idx, finfo in enumerate(filters_info, 1):
            f = finfo["filter"]
            f_type = finfo.get("type") or "-"
            cat_names = finfo.get("categories") or []
            has_override = finfo.get("has_override", False)
            visible = finfo.get("visible", True)

            try:
                fname = f.Name
            except:
                fname = u"<без имени>"

            try:
                link = output.linkify(f.Id, title=fname)
            except:
                link = fname

            cat_cell = u", ".join(cat_names) if cat_names else u"—"
            override_cell = u"🎨" if has_override else u"—"

            if visible:
                vis_cell = u"👁️"
            else:
                vis_cell = u"❌"

            filters_table_data.append([idx, link, f_type, cat_cell, override_cell, vis_cell])
        
        output.print_table(
            filters_table_data,
            columns=[u"№", u"Фильтр", u"Тип", u"Категории фильтра", u"Графика", u"Видимость"]
        )
        output.print_md(u"")
    else:
        output.print_md(u"### ⚠️ Для активного вида не найдено фильтров, которые бы отбирали этот элемент.")
        output.print_md(u"")

    # --- Дополнительная информация о видимости ---
    output.print_md(u"### 👁 Дополнительные настройки вида, влияющие на элемент")

    settings_table_data = []
    
    cat_name = visibility_info.get("category_name")
    settings_table_data.append([u"Категория элемента", cat_name or u"—"])

    if cat_hidden is None:
        cat_hidden_str = u"неизвестно"
    else:
        cat_hidden_str = u"скрыта" if cat_hidden else u"отображается"
    settings_table_data.append([u"Категория на виде", cat_hidden_str])

    cat_has_over = visibility_info.get("category_has_overrides", False)
    settings_table_data.append([u"Переопределения графики категории", u"есть" if cat_has_over else u"нет"])

    # Для элементов из связи - показываем статус категории "Связанные файлы"
    if linked_info:
        link_cat_name_show = visibility_info.get("link_category_name", u"Связанные файлы")
        if link_cat_hidden is None:
            link_cat_hidden_str = u"неизвестно"
        else:
            link_cat_hidden_str = u"скрыта" if link_cat_hidden else u"отображается"
        settings_table_data.append([u"Категория «{}»".format(link_cat_name_show), link_cat_hidden_str])

    elem_has_over = visibility_info.get("element_has_overrides", False)
    settings_table_data.append([u"Переопределения графики элемента", u"есть" if elem_has_over else u"нет"])

    # Рабочие наборы
    if workset_name:
        ws_open_str = u"открыт" if workset_is_open else u"закрыт"
        if workset_visible is None:
            ws_vis_str = u"неизвестно"
        elif workset_visible:
            ws_vis_str = u"отображается"
        else:
            ws_vis_str = u"скрыт"
        settings_table_data.append([u"Рабочий набор элемента", u"{} ({}, на виде: {})".format(workset_name, ws_open_str, ws_vis_str)])
    
    # Рабочий набор связи
    if link_workset_name:
        link_ws_open_str = u"открыт" if link_workset_is_open else u"закрыт"
        if link_workset_visible is None:
            link_ws_vis_str = u"неизвестно"
        elif link_workset_visible:
            link_ws_vis_str = u"отображается"
        else:
            link_ws_vis_str = u"скрыт"
        settings_table_data.append([u"Рабочий набор связи", u"{} ({}, на виде: {})".format(link_workset_name, link_ws_open_str, link_ws_vis_str)])

    # Объём подрезки вида
    if inside_crop is None:
        crop_str = u"не определено (подрезка отключена или данные недоступны)"
    else:
        crop_str = u"попадает в объём подрезки" if inside_crop else u"не попадает в объём подрезки"
    settings_table_data.append([u"Объём подрезки / секущий диапазон", crop_str])

    if tvm:
        if reveal_mode:
            tvm_str = u"RevealHiddenElements (Показать скрытые элементы)"
        else:
            tvm_str = tvm
        settings_table_data.append([u"Временный режим вида", tvm_str])
    else:
        settings_table_data.append([u"Временный режим вида", u"—"])

    output.print_table(settings_table_data, columns=[u"Параметр", u"Значение"])
    output.print_md(u"")
    output.print_md(u"---")
    output.print_md(u"_Подсказка: клик по имени фильтра откроет его свойства в Revit._")


def main():
    view = doc.ActiveView
    if view is None:
        TaskDialog.Show(
            "Проверка фильтров",
            "Нет активного вида."
        )
        return

    # Игнорируем спецификации и листы
    try:
        if view.ViewType in (ViewType.Schedule, ViewType.DrawingSheet, ViewType.Report):
            TaskDialog.Show(
                "Проверка фильтров",
                "Скрипт работает только на графических видах, а не на листах и спецификациях."
            )
            return
    except:
        pass

    element, linked_info = get_target_element()
    if element is None:
        TaskDialog.Show(
            "Проверка фильтров",
            "Не удалось получить элемент. Выберите один элемент и запустите скрипт снова."
        )
        return

    filters_info = find_filters_for_element(view, element, linked_info)
    visibility_info = get_visibility_info(view, element, linked_info)
    show_result(filters_info, visibility_info, view, element, linked_info)


if __name__ == "__main__":
    main()
