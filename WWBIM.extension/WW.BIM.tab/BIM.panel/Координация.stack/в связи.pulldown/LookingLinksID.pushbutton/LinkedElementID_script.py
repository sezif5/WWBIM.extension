# -*- coding: utf-8 -*-
# pyRevit button script: подрезка активного вида по ID элемента в выбранной связи
# - Всегда показывает список только ЗАГРУЖЕННЫХ связей и просит выбрать,
#   в какой связи искать элемент по ID.
# - Подрезает активный план, разрез, фасад или 3D-вид.
# - Для остальных видов использует персональный 3D-вид или создаёт новый.
# - В Revit 2023+ пытается подсветить КОНКРЕТНЫЙ элемент в связи через
#   Selection.SetReferences + Reference.CreateLinkReference.
#   В более старых версиях API выделяется только экземпляр связи.

import math

import clr
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ElementId,
    View3D,
    ViewPlan,
    ViewType,
    ViewFamily,
    ViewFamilyType,
    PlanViewPlane,
    PlanViewRange,
    Transaction,
    TransactionStatus,
    BoundingBoxXYZ,
    XYZ,
    Reference,
)
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import Clipboard
from System.Collections.Generic import List as CsList

from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

PAD = 3.0


def _show(title, text):
    try:
        TaskDialog.Show(title, text)
    except:
        pass


def _get_id_from_user():
    """Получить целочисленный ID элемента.
    Пытаемся взять текст из буфера обмена как значение по умолчанию,
    затем просим пользователя подтвердить/изменить его."""
    clip_text = None
    try:
        if Clipboard.ContainsText():
            clip_text = Clipboard.GetText()
    except:
        clip_text = None

    default = u""
    if clip_text:
        default = clip_text.strip()

    id_text = None
    try:
        id_text = forms.ask_for_string(
            prompt=u"Введите ID элемента из связанного файла (можно просто нажать Ctrl+V)",
            default=default,
            title=u"3D по ID элемента в связи",
        )
    except:
        _show(u"Ввод ID", u"Не удалось открыть форму ввода ID.")
        return None

    if id_text is None:
        return None

    id_text = id_text.strip()
    if not id_text:
        return None

    try:
        return int(id_text)
    except:
        _show(u"Некорректный ID", u"Не удалось преобразовать '{0}' в целое число.".format(id_text))
        return None


def _choose_link_instance():
    """Показать список всех ЗАГРУЖЕННЫХ связей и вернуть выбранный RevitLinkInstance."""
    links = []
    labels = []

    col = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    for link_inst in col:
        # Используем только загруженные связи (у которых есть LinkDocument)
        linked_doc = None
        try:
            linked_doc = link_inst.GetLinkDocument()
        except:
            linked_doc = None

        if linked_doc is None:
            continue

        links.append(link_inst)

        try:
            link_name = link_inst.Name
        except:
            link_name = u"(без имени)"

        try:
            doc_title = linked_doc.Title
        except:
            doc_title = u""

        label = u"{0} | {1}".format(link_name, doc_title)
        labels.append(label)

    if not links:
        _show(u"Связи", u"В документе нет загруженных связей.")
        return None

    if len(links) == 1:
        # одна загруженная связь — выбираем её автоматически
        return links[0]

    choice = forms.SelectFromList.show(
        labels,
        title=u"Выбор связи для поиска ID",
        button_name=u"Искать в этой связи",
        multiselect=False,
    )

    if not choice:
        return None

    idx = labels.index(choice)
    return links[idx]


def _find_element_in_links(element_int_id):
    """Найти элемент с указанным ID в выбранной пользователем связи.
    Возвращает (RevitLinkInstance, Element) или (None, None)."""
    target_id = ElementId(element_int_id)

    link_inst = _choose_link_instance()
    if link_inst is None:
        return (None, None)

    linked_doc = None
    try:
        linked_doc = link_inst.GetLinkDocument()
    except:
        linked_doc = None

    if linked_doc is None:
        _show(u"Связь", u"Не удалось получить документ выбранной связи.")
        return (None, None)

    try:
        linked_el = linked_doc.GetElement(target_id)
    except:
        linked_el = None

    if linked_el is None:
        _show(
            u"Поиск элемента",
            u"Элемент с ID {0} не найден в выбранной связи.".format(element_int_id),
        )
        return (None, None)

    return (link_inst, linked_el)


def _get_3d_view_family_type_id():
    vft_col = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    for vft in vft_col:
        try:
            if vft.ViewFamily == ViewFamily.ThreeDimensional:
                return vft.Id
        except:
            continue
    return None


def _get_box_corners(box):
    try:
        box_min = box.Min
        box_max = box.Max
        box_transform = box.Transform
    except:
        return []

    corners = []
    for x in (box_min.X, box_max.X):
        for y in (box_min.Y, box_max.Y):
            for z in (box_min.Z, box_max.Z):
                try:
                    corners.append(box_transform.OfPoint(XYZ(x, y, z)))
                except:
                    return []
    return corners


def _get_element_points_in_host(link_inst, linked_el):
    try:
        linked_box = linked_el.get_BoundingBox(None)
    except:
        linked_box = None

    if linked_box is None:
        _show(u"Подрезка вида", u"Не удалось получить границы элемента в связи.")
        return []

    try:
        link_transform = link_inst.GetTotalTransform()
    except:
        link_transform = None

    if link_transform is None:
        _show(u"Подрезка вида", u"Не удалось получить трансформацию связи.")
        return []

    host_points = []
    for point in _get_box_corners(linked_box):
        try:
            host_points.append(link_transform.OfPoint(point))
        except:
            return []
    return host_points


def _get_bounds_in_box_coordinates(points, box_transform):
    try:
        inverse = box_transform.Inverse
    except:
        return None

    local_points = []
    for point in points:
        try:
            local_points.append(inverse.OfPoint(point))
        except:
            return None

    if not local_points:
        return None

    return (
        min([point.X for point in local_points]),
        min([point.Y for point in local_points]),
        min([point.Z for point in local_points]),
        max([point.X for point in local_points]),
        max([point.Y for point in local_points]),
        max([point.Z for point in local_points]),
    )


def _create_box_for_points(points, source_box, keep_z=False):
    try:
        source_transform = source_box.Transform
        source_min = source_box.Min
        source_max = source_box.Max
    except:
        return None

    bounds = _get_bounds_in_box_coordinates(points, source_transform)
    if bounds is None:
        return None

    min_x, min_y, min_z, max_x, max_y, max_z = bounds
    if keep_z:
        min_z = source_min.Z
        max_z = source_max.Z

    result = BoundingBoxXYZ()
    try:
        result.Transform = source_transform
        result.Min = XYZ(min_x - PAD, min_y - PAD, min_z if keep_z else min_z - PAD)
        result.Max = XYZ(max_x + PAD, max_y + PAD, max_z if keep_z else max_z + PAD)
    except:
        return None
    return result


def _get_existing_or_personal_3d_view():
    """Вернуть уже открытый 3D-вид (если активен) или персональный {3D - Username}.
    Если ни один не найден, вернуть None (в этом случае создадим новый)."""
    # 1) Если активный вид — 3D и не шаблон, используем его
    av = uidoc.ActiveView
    try:
        if isinstance(av, View3D) and not av.IsTemplate:
            return av
    except:
        pass

    # 2) Ищем персональный 3D-вид пользователя вида {3D - Username}
    try:
        username = doc.Application.Username
    except:
        username = None

    if username:
        target_name = u"{3D - %s}" % username
        v_col = FilteredElementCollector(doc).OfClass(View3D)
        for v in v_col:
            try:
                if (not v.IsTemplate) and v.Name == target_name:
                    return v
            except:
                continue

    return None


def _prepare_3d_view_with_section_box(points):
    view3d = None
    t = Transaction(doc, u"Подрезка 3D по элементу в связи")
    try:
        t.Start()
    except:
        _show(u"3D вид", u"Не удалось начать изменение 3D-вида.")
        return None

    try:
        view3d = _get_existing_or_personal_3d_view()
        if view3d is None:
            vft_id = _get_3d_view_family_type_id()
            if vft_id is None:
                t.RollBack()
                _show(u"3D вид", u"В документе не найден тип 3D-вида.")
                return None
            view3d = View3D.CreateIsometric(doc, vft_id)

        try:
            source_box = view3d.GetSectionBox()
        except:
            try:
                source_box = view3d.SectionBox
            except:
                source_box = None

        bbox = _create_box_for_points(points, source_box, keep_z=False)
        if bbox is None:
            t.RollBack()
            _show(u"3D вид", u"Не удалось рассчитать секционный бокс.")
            return None

        try:
            view3d.SetSectionBox(bbox)
        except:
            view3d.SectionBox = bbox

        try:
            view3d.IsSectionBoxActive = True
        except:
            pass

        status = t.Commit()
        if status == TransactionStatus.Pending:
            _show(u"3D вид", u"Revit ожидает обработки предупреждений транзакции.")
            return None
        if status != TransactionStatus.Committed:
            _show(u"3D вид", u"Изменение 3D-вида не было зафиксировано.")
            return None
    except Exception as error:
        try:
            t.RollBack()
        except:
            pass
        _show(u"3D вид", u"Не удалось подрезать 3D-вид:\n{0}".format(error))
        return None

    return view3d


def _is_supported_crop_view(view):
    try:
        if view.IsTemplate:
            return False
    except:
        return False

    try:
        if isinstance(view, ViewPlan):
            return True
    except:
        pass

    try:
        return view.ViewType in (ViewType.Section, ViewType.Elevation)
    except:
        return False


def _has_custom_crop_shape(view):
    try:
        manager = view.GetCropRegionShapeManager()
    except:
        return False

    try:
        if manager.ShapeSet:
            return True
    except:
        pass

    try:
        if manager.NumberOfSplitRegions > 1:
            return True
    except:
        pass

    return False


def _get_range_plane_data(view, view_range, plane):
    try:
        level_id = view_range.GetLevelId(plane)
    except:
        return None

    try:
        if level_id == PlanViewRange.Unlimited:
            if plane == PlanViewPlane.TopClipPlane:
                return (None, float("inf"))
            return (None, float("-inf"))
    except:
        pass

    try:
        level = doc.GetElement(level_id)
    except:
        level = None

    if level is None:
        return None

    try:
        base_elevation = level.ProjectElevation
    except:
        try:
            base_elevation = level.Elevation
        except:
            return None

    try:
        offset = view_range.GetOffset(plane)
    except:
        return None

    return (base_elevation, base_elevation + offset)


def _format_elevation(value):
    try:
        if math.isinf(value):
            return u"Без ограничений"
    except:
        pass
    return u"{0:+.3f} м".format(value * 0.3048)


def _ask_plan_range_change(view, points):
    try:
        view_range = view.GetViewRange()
    except:
        return None

    top_data = _get_range_plane_data(view, view_range, PlanViewPlane.TopClipPlane)
    bottom_data = _get_range_plane_data(view, view_range, PlanViewPlane.BottomClipPlane)
    depth_data = _get_range_plane_data(view, view_range, PlanViewPlane.ViewDepthPlane)
    if top_data is None or bottom_data is None or depth_data is None:
        _show(
            u"Секущий диапазон",
            u"Не удалось проверить секущий диапазон. Будет изменена только граница подрезки.",
        )
        return None

    element_min_z = min([point.Z for point in points])
    element_max_z = max([point.Z for point in points])
    old_top = top_data[1]
    old_bottom = bottom_data[1]
    old_depth = depth_data[1]
    visible_bottom = min(old_bottom, old_depth)
    if element_max_z >= visible_bottom and element_min_z <= old_top:
        return None

    new_top = max(old_top, element_max_z + PAD)
    new_bottom = min(old_bottom, element_min_z - PAD)
    new_depth = min(old_depth, new_bottom)

    message = (
        u"Элемент не попадает в текущий секущий диапазон.\n\n"
        u"Текущий диапазон:\n"
        u"Верх: {0}\nНиз: {1}\nГлубина: {2}\n\n"
        u"Новый диапазон:\n"
        u"Верх: {3}\nНиз: {4}\nГлубина: {5}\n\n"
        u"Изменить секущий диапазон?\n"
        u"При выборе «Нет» изменится только граница подрезки, и элемент может остаться невидимым."
    ).format(
        _format_elevation(old_top),
        _format_elevation(old_bottom),
        _format_elevation(old_depth),
        _format_elevation(new_top),
        _format_elevation(new_bottom),
        _format_elevation(new_depth),
    )

    try:
        accepted = forms.alert(
            message,
            title=u"Секущий диапазон",
            yes=True,
            no=True,
            warn_icon=True,
        )
    except:
        accepted = False

    if not accepted:
        return None

    changes = []
    for plane, plane_data, new_elevation in (
        (PlanViewPlane.TopClipPlane, top_data, new_top),
        (PlanViewPlane.BottomClipPlane, bottom_data, new_bottom),
        (PlanViewPlane.ViewDepthPlane, depth_data, new_depth),
    ):
        base_elevation, old_elevation = plane_data
        if base_elevation is None or abs(new_elevation - old_elevation) < 0.000001:
            continue
        changes.append((plane, base_elevation, new_elevation))

    return (view_range, changes) if changes else None


def _crop_active_view(view, points, range_change):
    try:
        source_box = view.CropBox
    except:
        source_box = None

    try:
        is_plan = isinstance(view, ViewPlan)
    except:
        is_plan = False

    crop_box = _create_box_for_points(points, source_box, keep_z=is_plan)
    if crop_box is None:
        return False

    t = Transaction(doc, u"Подрезка вида по элементу в связи")
    try:
        t.Start()
    except:
        return False

    try:
        if range_change is not None:
            view_range, changes = range_change
            for plane, base_elevation, new_elevation in changes:
                view_range.SetOffset(plane, new_elevation - base_elevation)
            view.SetViewRange(view_range)

        view.CropBox = crop_box
        view.CropBoxActive = True
        status = t.Commit()
        if status == TransactionStatus.Pending:
            return None
        return status == TransactionStatus.Committed
    except:
        try:
            t.RollBack()
        except:
            pass
        return False


def _prepare_view_with_crop(link_inst, linked_el):
    points = _get_element_points_in_host(link_inst, linked_el)
    if not points:
        return None

    try:
        active_view = uidoc.ActiveView
    except:
        active_view = None

    try:
        if isinstance(active_view, View3D):
            return _prepare_3d_view_with_section_box(points)
    except:
        pass

    if active_view is not None and _is_supported_crop_view(active_view):
        if _has_custom_crop_shape(active_view):
            _show(
                u"Фигурная подрезка",
                u"Активный вид имеет фигурную или разделённую подрезку. Она не будет изменена; элемент будет открыт в 3D.",
            )
        else:
            try:
                is_plan = isinstance(active_view, ViewPlan)
            except:
                is_plan = False
            range_change = _ask_plan_range_change(active_view, points) if is_plan else None
            crop_result = _crop_active_view(active_view, points, range_change)
            if crop_result is True:
                return active_view
            if crop_result is None:
                _show(
                    u"Подрезка вида",
                    u"Revit ожидает обработки предупреждений транзакции. Повторите команду после их закрытия.",
                )
                return None
            _show(
                u"Подрезка вида",
                u"Не удалось изменить подрезку активного вида. Элемент будет открыт в 3D.",
            )

    return _prepare_3d_view_with_section_box(points)


def _select_link_or_element_in_view(target_view, link_inst, linked_el):
    """Делаем целевой вид активным и по возможности выделяем КОНКРЕТНЫЙ элемент
    в связи (через Selection.SetReferences + Reference.CreateLinkReference).
    Если API этого не поддерживает, просто выделяем экземпляр связи."""
    if target_view is None:
        return

    try:
        uidoc.ActiveView = target_view
    except:
        return

    selection = uidoc.Selection

    # Пытаемся использовать Selection.SetReferences (Revit 2023+)
    has_set_refs = False
    try:
        # hasattr на .NET-метод в IronPython работает
        has_set_refs = hasattr(selection, 'SetReferences')
    except:
        has_set_refs = False

    if has_set_refs:
        try:
            # Reference на элемент в связке
            ref_in_link = Reference(linked_el)
            # Преобразуем в reference в хосте для конкретного экземпляра связи
            ref_in_host = ref_in_link.CreateLinkReference(link_inst)

            from System.Collections.Generic import List as CsRefList
            ref_list = CsRefList[Reference]()
            ref_list.Add(ref_in_host)

            selection.SetReferences(ref_list)
            uidoc.ShowElements(ref_in_host)
            return
        except:
            # Если что-то пошло не так — падаем в запасной вариант
            pass

    # Fallback: просто выделяем экземпляр связи (как минимум не падаем)
    try:
        ids = CsList[ElementId]()
        ids.Add(link_inst.Id)
        selection.SetElementIds(ids)
        uidoc.ShowElements(link_inst.Id)
    except:
        pass


def main():
    element_int_id = _get_id_from_user()
    if element_int_id is None:
        return

    link_inst, linked_el = _find_element_in_links(element_int_id)
    if link_inst is None or linked_el is None:
        return

    target_view = _prepare_view_with_crop(link_inst, linked_el)
    _select_link_or_element_in_view(target_view, link_inst, linked_el)


if __name__ == "__main__":
    main()
