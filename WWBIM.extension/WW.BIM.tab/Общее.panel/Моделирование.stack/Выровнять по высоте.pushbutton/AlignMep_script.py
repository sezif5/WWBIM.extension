# coding: utf-8
# AlignMep_script
# Выровнять по высоте (MEP) v6
# ----------------------------------------
# Скрипт для pyRevit: выравнивание одной MEP-кривой по высоте относительно другой
# с учётом уклона и для пересекающихся, параллельных и последовательных участков.
#
# Логика:
#   1. Пользователь выбирает базовую MEP-кривую (магистраль).
#   2. Пользователь выбирает вторую MEP-кривую (ветку), которую нужно поднять/опустить.
#   3. Если оси труб в плане пересекаются — считаем высоту в точке пересечения
#      (учитывая уклон обеих труб) и выравниваем по ней.
#   4. Если оси параллельны:
#        4.1. Если ветка лежит на той же оси (последовательные/соосные участки),
#             считаем высоту в предполагаемой точке стыка на оси (по базовой трубе)
#             и выравниваем так, чтобы уклон сохранялся и стык был корректен.
#        4.2. Если оси параллельны, но смещены в плане — используется проекция
#             ближайшего конца ветки на базовую ось (как в v4).
#
# В результате в предполагаемом месте подключения трубы имеют одну и ту же отметку.
#
# Ограничения:
#   - Работает для MEPCurve (Pipe, Duct, CableTray и т.п.) с линейной геометрией.
#   - Уклон не меняется, выполняется только вертикальное смещение ветки.
#
# pyRevit:
#   Поместите файл как скрипт кнопки, запускайте из контекста проекта.

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


class MEPCurveSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, MEPCurve)

    def AllowReference(self, reference, position):
        return True


def pick_mepcurve(prompt):
    sel_filter = MEPCurveSelectionFilter()
    ref = uidoc.Selection.PickObject(ObjectType.Element, sel_filter, prompt)
    return doc.GetElement(ref.ElementId)


def get_endpoints(mepcurve):
    loc = mepcurve.Location
    loc_curve = getattr(loc, 'Curve', None)
    if not loc_curve:
        return None, None, None
    p0 = loc_curve.GetEndPoint(0)
    p1 = loc_curve.GetEndPoint(1)
    return loc_curve, p0, p1


def format_delta_mm(dz_internal):
    try:
        mm = UnitUtils.ConvertFromInternalUnits(abs(dz_internal), DisplayUnitType.DUT_MILLIMETERS)
        return u'{0:.1f} мм'.format(mm)
    except:
        return u'{0:.4f} футов'.format(abs(dz_internal))


def xy_intersection_params(p1, p2, p3, p4, tol=1e-9):
    """По двум отрезкам (p1-p2 и p3-p4) находит пересечение бесконечных прямых
    в плоскости XY. Возвращает (t1, t2), параметры вдоль первого и второго отрезка,
    такие что:
        P = p1 + t1 * (p2 - p1) = p3 + t2 * (p4 - p3)  (по XY)
    Если прямые параллельны в плане, возвращает (None, None)."""
    v1x = p2.X - p1.X
    v1y = p2.Y - p1.Y
    v2x = p4.X - p3.X
    v2y = p4.Y - p3.Y

    dx = p3.X - p1.X
    dy = p3.Y - p1.Y

    a = v1x
    b = -v2x
    c = v1y
    d = -v2y

    det = a * d - b * c
    if abs(det) < tol:
        return None, None

    t1 = (dx * d - b * dy) / det
    t2 = (a * dy - dx * c) / det
    return t1, t2


def point_on_segment(p_start, p_end, t):
    return XYZ(
        p_start.X + t * (p_end.X - p_start.X),
        p_start.Y + t * (p_end.Y - p_start.Y),
        p_start.Z + t * (p_end.Z - p_start.Z)
    )


def project_endpoint_mode(base_curve, m0, m1):
    """Режим для параллельных, но смещённых осей (как v4):
    берётся конечная точка ветки, ближайшая к базовой кривой."""
    move_endpoints = [m0, m1]
    best_move_pt = None
    best_base_pt = None
    best_dist = None

    for mp in move_endpoints:
        try:
            proj_result = base_curve.Project(mp)
        except:
            proj_result = None

        if proj_result is None:
            continue

        base_pt = proj_result.XYZPoint
        dist = proj_result.Distance

        if best_dist is None or dist < best_dist:
            best_dist = dist
            best_move_pt = mp
            best_base_pt = base_pt

    return best_move_pt, best_base_pt


def handle_parallel_case(base_curve, b0, b1, m0, m1, tol_axis=1e-6):
    """Обработка параллельных труб.
    Если ветка лежит на той же оси (соосные/последовательные участки),
    считаем высоту в точке стыка по оси базовой трубы.
    Иначе используем project_endpoint_mode (смещённые параллельные)."""
    # Вектор оси базовой и ветки в XY
    vbx = b1.X - b0.X
    vby = b1.Y - b0.Y
    vmx = m1.X - m0.X
    vmy = m1.Y - m0.Y

    # Проверяем, лежит ли m0 на оси базовой (расстояние до прямой в XY)
    # Нормаль к оси (vbx, vby): (-vby, vbx)
    num = abs((m0.X - b0.X) * vby - (m0.Y - b0.Y) * vbx)
    den = (vbx * vbx + vby * vby) ** 0.5
    dist_to_axis = num / den if den > 0 else 0.0

    if dist_to_axis < tol_axis:
        # Соосные участки: предполагаем стык в районе ближайших концов
        candidates = [
            (b0, m0),
            (b0, m1),
            (b1, m0),
            (b1, m1)
        ]

        best_pair = None
        best_d = None
        for bp, mp in candidates:
            dx = bp.X - mp.X
            dy = bp.Y - mp.Y
            d = (dx * dx + dy * dy) ** 0.5
            if best_d is None or d < best_d:
                best_d = d
                best_pair = (bp, mp)

        base_conn, move_conn = best_pair

        # Находим точку на ветке, которая имеет те же XY, что и base_conn,
        # вдоль её оси (расширяем отрезок при необходимости).
        vmx3 = m1.X - m0.X
        vmy3 = m1.Y - m0.Y
        vmz3 = m1.Z - m0.Z

        denom2 = vmx3 * vmx3 + vmy3 * vmy3
        if denom2 == 0.0:
            # Деградировавший случай (нулевая длина) — fallback
            return project_endpoint_mode(base_curve, m0, m1)

        t_move = ((base_conn.X - m0.X) * vmx3 + (base_conn.Y - m0.Y) * vmy3) / denom2
        move_pt = point_on_segment(m0, m1, t_move)
        base_pt = base_conn
        return move_pt, base_pt

    # Не соосные, просто параллельные оси
    return project_endpoint_mode(base_curve, m0, m1)


def main():
    try:
        base_curve_el = pick_mepcurve(u'Выберите базовую MEP-кривую (магистраль)')
        move_curve_el = pick_mepcurve(u'Выберите MEP-кривую, которую нужно выровнять по высоте')
    except Exception:
        TaskDialog.Show(u'Выровнять по высоте', u'Операция отменена пользователем.')
        return

    if base_curve_el.Id == move_curve_el.Id:
        TaskDialog.Show(u'Выровнять по высоте', u'Выбраны дважды один и тот же элемент. Выберите две разные MEP-кривые.')
        return

    base_curve, b0, b1 = get_endpoints(base_curve_el)
    move_curve, m0, m1 = get_endpoints(move_curve_el)

    if not base_curve or not move_curve:
        TaskDialog.Show(u'Выровнять по высоте', u'Не удалось получить геометрию одной из кривых.')
        return

    # Пытаемся найти пересечение осей труб в плане
    t1, t2 = xy_intersection_params(b0, b1, m0, m1)

    if t1 is not None and t2 is not None:
        # Есть пересечение в плане: считаем отметки на обеих трубах в этой XY-точке
        base_pt = point_on_segment(b0, b1, t1)
        move_pt = point_on_segment(m0, m1, t2)
    else:
        # Параллельные оси: обрабатываем соосный / параллельный случай
        move_pt, base_pt = handle_parallel_case(base_curve, b0, b1, m0, m1)
        if move_pt is None or base_pt is None:
            TaskDialog.Show(u'Выровнять по высоте', u'Не удалось определить рабочую точку выравнивания.')
            return

    dz = base_pt.Z - move_pt.Z

    tolerance_ft = 0.0003  # ~0.09 мм
    if abs(dz) < tolerance_ft:
        msg = u'Кривые уже выровнены по высоте.\nРазница по Z: {0}'.format(format_delta_mm(dz))
        TaskDialog.Show(u'Выровнять по высоте', msg)
        return

    move_vec = XYZ(0.0, 0.0, dz)

    t = Transaction(doc, u'Выровнять по высоте MEP-кривой')
    t.Start()
    try:
        ElementTransformUtils.MoveElement(doc, move_curve_el.Id, move_vec)
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show(u'Выровнять по высоте', u'Ошибка при выравнивании:\n{0}'.format(str(e)))


if __name__ == '__main__':
    main()
