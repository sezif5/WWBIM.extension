# -*- coding: utf-8 -*-
from __future__ import print_function

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import BuiltInCategory, Plane, SketchPlane


def selected_room(doc, uidoc):
    rooms = []
    for element_id in uidoc.Selection.GetElementIds():
        try:
            element = doc.GetElement(element_id)
            if element is not None and element.Category is not None and \
                    element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms):
                rooms.append(element)
        except Exception:
            continue
    if len(rooms) != 1:
        return None, len(rooms)
    return rooms[0], 1


def curve_points(curve):
    points = []
    try:
        for point in curve.Tessellate():
            if not points or points[-1].DistanceTo(point) > 1e-7:
                points.append(point)
    except Exception:
        try:
            points = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
        except Exception:
            return []
    return points


def loop_area(loop):
    points = []
    for curve in loop:
        current = curve_points(curve)
        if not current:
            continue
        if points and points[-1].DistanceTo(current[0]) < 1e-7:
            points.extend(current[1:])
        else:
            points.extend(current)
    if len(points) < 3:
        return 0.0
    if points[0].DistanceTo(points[-1]) > 1e-7:
        points.append(points[0])
    area = 0.0
    for index in range(len(points) - 1):
        area += points[index].X * points[index + 1].Y
        area -= points[index + 1].X * points[index].Y
    return abs(area) * 0.5


def get_room_loops(room):
    options = DB.SpatialElementBoundaryOptions()
    boundary_lists = room.GetBoundarySegments(options)
    if not boundary_lists:
        return []
    loops = []
    for boundary_list in boundary_lists:
        curves = []
        for segment in boundary_list:
            try:
                curve = segment.GetCurve()
                if curve is not None and curve.IsBound:
                    curves.append(curve)
            except Exception:
                continue
        if curves and loop_area(curves) > 1e-8:
            loops.append((loop_area(curves), curves))
    return loops


def target_elevation(view, fallback):
    try:
        if view.GenLevel is not None:
            return view.GenLevel.Elevation
    except Exception:
        pass
    return fallback


def move_curve_to_level(curve, elevation):
    start = curve.GetEndPoint(0)
    return curve.CreateTransformed(
        DB.Transform.CreateTranslation(DB.XYZ(0, 0, elevation - start.Z)))


def room_name(room):
    try:
        name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
        number = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()
        return u"{0} {1}".format(number or u"", name or u"").strip()
    except Exception:
        return u"ID {0}".format(room.Id.IntegerValue)


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    view = revit.active_view
    try:
        if view is None or view.ViewType != DB.ViewType.AreaPlan or view.IsTemplate:
            forms.alert(u"Запустите инструмент с активного плана зонирования.",
                        title=u"Границы зон по помещению")
            return

        room, count = selected_room(doc, uidoc)
        if room is None:
            if count == 0:
                message = u"Выделите одно помещение, расположенное вокруг здания, и повторите запуск."
            else:
                message = u"Должно быть выделено ровно одно помещение. Сейчас выделено помещений: {0}.".format(count)
            forms.alert(message, title=u"Границы зон по помещению")
            return

        loops = get_room_loops(room)
        if not loops:
            forms.alert(u"У помещения «{0}» не найден замкнутый внутренний контур.".format(room_name(room)),
                        title=u"Границы зон по помещению")
            return

        loops.sort(key=lambda item: item[0])
        selected_area, selected_curves = loops[0]
        if len(loops) == 1:
            question = u"У помещения найден один контур площадью {0:.2f} м². Создать по нему границы зон?".format(
                selected_area * 0.092903)
        else:
            question = u"У помещения найдено контуров: {0}.\nБудет использован внутренний контур площадью {1:.2f} м².\n\nПродолжить?".format(
                len(loops), selected_area * 0.092903)
        if not forms.alert(question, title=u"Границы зон по помещению", yes=True, no=True):
            return

        try:
            source_z = selected_curves[0].GetEndPoint(0).Z
        except Exception:
            source_z = 0.0
        elevation = target_elevation(view, source_z)
        transaction = DB.Transaction(doc, u"Создать границы зон по помещению")
        created = 0
        failed = 0
        try:
            transaction.Start()
            sketch_plane = SketchPlane.Create(
                doc, Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elevation)))
            for source_curve in selected_curves:
                try:
                    curve = move_curve_to_level(source_curve, elevation)
                    boundary = doc.Create.NewAreaBoundaryLine(sketch_plane, curve, view)
                    if boundary is not None:
                        created += 1
                    else:
                        failed += 1
                except Exception:
                    failed += 1
            if created == 0:
                transaction.RollBack()
                forms.alert(u"Не удалось создать границы зон. Существующие линии не изменены.",
                            title=u"Границы зон по помещению")
                return
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise

        message = u"Создано границ зон: {0}.\nПомещение: {1}.\nСуществующие линии не удалялись.".format(
            created, room_name(room))
        if failed:
            message += u"\nНе удалось создать сегментов: {0}.".format(failed)
        forms.alert(message, title=u"Границы зон по помещению")
    except Exception as error:
        forms.alert(u"Ошибка создания границ зон:\n{0}".format(error),
                    title=u"Границы зон по помещению")


if __name__ == u"__main__":
    main()
