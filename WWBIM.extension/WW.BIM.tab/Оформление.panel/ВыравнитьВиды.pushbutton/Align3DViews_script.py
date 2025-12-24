
# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script

__doc__ = u"Выравнивание 3D-видов на листе: совмещает виды так, чтобы одна и та же точка модели была в одном месте на листе."

logger = script.get_logger()
UIDOC = revit.uidoc
DOC = revit.doc


def _get_selected_viewports(doc, uidoc):
    """Получить два выбранных viewport'а на активном листе."""
    av = doc.ActiveView
    if not isinstance(av, DB.ViewSheet):
        forms.alert(
            u"Команда работает только на виде листа.",
            title=u"Выравнивание 3D-видов",
            warn_icon=True
        )
        return None, None

    sel_ids = list(uidoc.Selection.GetElementIds())
    vports = [doc.GetElement(eid) for eid in sel_ids if isinstance(doc.GetElement(eid), DB.Viewport)]

    if len(vports) != 2:
        forms.alert(
            u"Выберите на листе ровно ДВА viewport'а (сначала базовый, затем тот, который нужно подтянуть) и запустите команду ещё раз.",
            title=u"Выравнивание 3D-видов",
            warn_icon=True
        )
        return None, None

    base_vp, moved_vp = vports[0], vports[1]
    return base_vp, moved_vp


def _get_view_from_viewport(doc, viewport):
    if viewport is None:
        return None
    return doc.GetElement(viewport.ViewId)


def _project_point_to_view_plane(view3d, model_point):
    """
    Проецирует точку модели на плоскость вида.
    Возвращает (x, y) в координатах вида (right, up).
    """
    orientation = view3d.GetOrientation()
    eye = orientation.EyePosition
    forward = orientation.ForwardDirection
    up = orientation.UpDirection
    right = forward.CrossProduct(up)
    
    # Нормализуем
    try:
        forward = forward.Normalize()
        up = up.Normalize()
        right = right.Normalize()
    except:
        pass
    
    # Вектор от глаза к точке
    v = model_point - eye
    
    # Проекция на оси вида
    x = v.DotProduct(right)
    y = v.DotProduct(up)
    
    return x, y


def _get_view_projection_on_sheet(view3d, viewport, model_point):
    """
    Вычисляет позицию на листе, где отображается данная точка модели.
    
    Логика:
    1. Проецируем точку на плоскость вида (получаем x, y в координатах вида)
    2. Проецируем центр section box на плоскость вида (это соответствует центру viewport)
    3. Вычисляем разницу и переводим в координаты листа с учётом масштаба
    """
    # Получаем section box
    try:
        bb = view3d.GetSectionBox()
    except:
        bb = None
    
    if bb is None or not view3d.IsSectionBoxActive:
        bb = view3d.CropBox
    
    if bb is None:
        return None
    
    # Центр section box в мировых координатах
    t = bb.Transform
    local_center = (bb.Min + bb.Max) * 0.5
    world_center = t.OfPoint(local_center)
    
    # Проецируем нашу точку и центр box на плоскость вида
    pt_x, pt_y = _project_point_to_view_plane(view3d, model_point)
    center_x, center_y = _project_point_to_view_plane(view3d, world_center)
    
    # Смещение точки относительно центра section box (в модельных единицах)
    dx_model = pt_x - center_x
    dy_model = pt_y - center_y
    
    # Масштаб вида
    view_scale = view3d.Scale
    
    # Переводим смещение в координаты листа
    dx_sheet = dx_model / view_scale
    dy_sheet = dy_model / view_scale
    
    # Получаем центр СОДЕРЖИМОГО viewport'а (без учёта label'а)
    # Используем GetLabelOutline и GetBoxOutline для вычисления реального центра вида
    try:
        # Пробуем получить смещение label'а
        label_offset = viewport.LabelOffset
        label_line_length = viewport.LabelLineLength
    except:
        label_offset = DB.XYZ(0, 0, 0)
        label_line_length = 0
    
    # GetBoxCenter включает label, нам нужен центр самого вида
    box_center = viewport.GetBoxCenter()
    
    # Пробуем получить outline самого вида (без label)
    try:
        # В Revit 2023+ есть метод GetLabelOutline
        label_outline = viewport.GetLabelOutline()
        box_outline = viewport.GetBoxOutline()
        
        # Вычисляем центр только области вида (исключая label)
        # Label обычно снизу или сверху
        label_height = label_outline.MaximumPoint.Y - label_outline.MinimumPoint.Y
        
        # Проверяем, где находится label (сверху или снизу)
        if label_outline.MinimumPoint.Y < box_outline.MinimumPoint.Y + 0.001:
            # Label снизу — сдвигаем центр вверх
            view_center_y = box_center.Y + label_height / 2.0
        elif label_outline.MaximumPoint.Y > box_outline.MaximumPoint.Y - 0.001:
            # Label сверху — сдвигаем центр вниз
            view_center_y = box_center.Y - label_height / 2.0
        else:
            view_center_y = box_center.Y
            
        view_center_x = box_center.X
    except:
        # Если не удалось получить label outline, используем box center как есть
        view_center_x = box_center.X
        view_center_y = box_center.Y
    
    sheet_x = view_center_x + dx_sheet
    sheet_y = view_center_y + dy_sheet
    
    return DB.XYZ(sheet_x, sheet_y, 0.0)


def _get_reference_point(base_view, moved_view):
    """
    Определить опорную точку в модели для выравнивания.
    Лучше всего использовать начало координат (0,0,0) — это то, 
    что Revit показывает пунктирными линиями при перетаскивании.
    """
    # Используем начало координат — это даёт тот же результат,
    # что и пунктирные линии Revit при перетаскивании viewport'а
    return DB.XYZ(0, 0, 0)


def main():
    base_vp, moved_vp = _get_selected_viewports(DOC, UIDOC)
    if not base_vp or not moved_vp:
        return

    base_view = _get_view_from_viewport(DOC, base_vp)
    moved_view = _get_view_from_viewport(DOC, moved_vp)

    if not isinstance(base_view, DB.View3D) or not isinstance(moved_view, DB.View3D):
        forms.alert(
            u"Оба выбранных viewport'а должны ссылаться на 3D-виды.",
            title=u"Выравнивание 3D-видов",
            warn_icon=True
        )
        return

    if base_view.IsPerspective or moved_view.IsPerspective:
        forms.alert(
            u"Перспективные виды не поддерживаются. Используйте ортогональные 3D-виды.",
            title=u"Выравнивание 3D-видов",
            warn_icon=True
        )
        return
    
    # Проверяем, что ориентация видов совпадает
    orient1 = base_view.GetOrientation()
    orient2 = moved_view.GetOrientation()
    
    fwd1 = orient1.ForwardDirection
    fwd2 = orient2.ForwardDirection
    up1 = orient1.UpDirection
    up2 = orient2.UpDirection
    
    # Проверяем параллельность направлений (допуск ~1 градус)
    dot_fwd = abs(fwd1.DotProduct(fwd2))
    dot_up = abs(up1.DotProduct(up2))
    
    if dot_fwd < 0.9998 or dot_up < 0.9998:
        forms.alert(
            u"Виды должны иметь одинаковую ориентацию камеры.\n"
            u"Поверните виды так, чтобы они смотрели в одном направлении.",
            title=u"Выравнивание 3D-видов",
            warn_icon=True
        )
        return

    # Опорная точка — начало координат (0,0,0)
    # Это та же точка, которую Revit показывает пунктирными линиями
    ref_point = _get_reference_point(base_view, moved_view)
    logger.debug(u"Опорная точка: ({:.2f}, {:.2f}, {:.2f})".format(ref_point.X, ref_point.Y, ref_point.Z))
    
    # Вычисляем, где эта точка находится на листе для каждого viewport
    base_sheet_pt = _get_view_projection_on_sheet(base_view, base_vp, ref_point)
    moved_sheet_pt = _get_view_projection_on_sheet(moved_view, moved_vp, ref_point)
    
    if base_sheet_pt is None or moved_sheet_pt is None:
        forms.alert(
            u"Не удалось вычислить проекцию опорной точки.\n"
            u"Убедитесь, что у видов включён Section Box.",
            title=u"Выравнивание 3D-видов",
            warn_icon=True
        )
        return
    
    logger.debug(u"Точка на листе (базовый): ({:.4f}, {:.4f})".format(base_sheet_pt.X, base_sheet_pt.Y))
    logger.debug(u"Точка на листе (перемещаемый): ({:.4f}, {:.4f})".format(moved_sheet_pt.X, moved_sheet_pt.Y))
    
    # Вычисляем смещение
    dx = base_sheet_pt.X - moved_sheet_pt.X
    dy = base_sheet_pt.Y - moved_sheet_pt.Y
    
    logger.debug(u"Смещение: dx={:.4f}, dy={:.4f}".format(dx, dy))

    if abs(dx) < 1e-6 and abs(dy) < 1e-6:
        forms.alert(
            u"Виды уже выровнены!",
            title=u"Выравнивание 3D-видов",
            warn_icon=False
        )
        return

    t = DB.Transaction(DOC, u"Выравнивание 3D-видов")
    t.Start()
    try:
        move_vec = DB.XYZ(dx, dy, 0.0)
        DB.ElementTransformUtils.MoveElement(DOC, moved_vp.Id, move_vec)
        t.Commit()
        forms.alert(
            u"Виды выровнены!\n"
            u"Смещение: {:.2f} мм по X, {:.2f} мм по Y".format(dx * 304.8, dy * 304.8),
            title=u"Выравнивание 3D-видов",
            warn_icon=False
        )
    except Exception as exc:
        t.RollBack()
        logger.error(u"Ошибка при смещении viewport'а: {0}".format(exc))
        forms.alert(
            u"Произошла ошибка при смещении второго вида.\nПодробности смотрите в журнале pyRevit.",
            title=u"Выравнивание 3D-видов",
            warn_icon=True
        )


if __name__ == '__main__':
    main()
