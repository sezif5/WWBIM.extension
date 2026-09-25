# -*- coding: utf-8 -*-
from __future__ import print_function

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import BuiltInCategory, Curve, Line, LocationCurve, Plane, SketchPlane, ViewPlan
from System.Collections.Generic import List


def view_label(view):
    try:
        area_scheme = view.AreaScheme
        scheme_name = area_scheme.Name if area_scheme is not None else u"без схемы"
    except Exception:
        scheme_name = u"без схемы"
    try:
        level = view.GenLevel
        level_name = level.Name if level is not None else u"без уровня"
    except Exception:
        level_name = u"без уровня"
    try:
        return u"План зонирования ({0}) — {1} — {2}".format(
            scheme_name, view.Name, level_name)
    except Exception:
        return u"ID {0}".format(view.Id.IntegerValue)


def selected_room_separators(doc, uidoc):
    result = []
    for element_id in uidoc.Selection.GetElementIds():
        try:
            element = doc.GetElement(element_id)
            if element is None or element.Category is None:
                continue
            if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_RoomSeparationLines):
                location = element.Location
                if isinstance(location, LocationCurve) and location.Curve is not None:
                    result.append(element)
        except Exception:
            continue
    return result


def target_area_plans(doc):
    result = []
    for view in DB.FilteredElementCollector(doc).OfClass(ViewPlan):
        try:
            if view.IsTemplate or view.ViewType != DB.ViewType.AreaPlan:
                continue
            result.append(view)
        except Exception:
            continue
    return sorted(result, key=lambda item: view_label(item).lower())


def target_elevation(view, fallback):
    try:
        if view.GenLevel is not None:
            return view.GenLevel.Elevation
    except Exception:
        pass
    return fallback


def copy_curve_to_elevation(curve, elevation):
    try:
        if curve is None or not curve.IsBound:
            return None
        start = curve.GetEndPoint(0)
        dz = elevation - start.Z
        transformed = curve.CreateTransformed(DB.Transform.CreateTranslation(DB.XYZ(0, 0, dz)))
        return transformed
    except Exception:
        return None


def create_area_boundary(doc, target_view, curve, sketch_plane):
    try:
        return doc.Create.NewAreaBoundaryLine(sketch_plane, curve, target_view)
    except Exception:
        return None


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    try:
        separators = selected_room_separators(doc, uidoc)
        if not separators:
            forms.alert(u"Сначала выделите элементы категории «Разделитель помещений» на плане.",
                        title=u"Разделители в границы зон")
            return

        views = target_area_plans(doc)
        if not views:
            forms.alert(u"В проекте не найдено ни одного плана площадей.",
                        title=u"Разделители в границы зон")
            return
        labels = [view_label(view) for view in views]
        selected_label = forms.SelectFromList.show(
            labels,
            title=u"Выберите план зонирования",
            button_name=u"Копировать границы",
            multiselect=False)
        if not selected_label:
            return
        target_view = views[labels.index(selected_label)]

        first_curve = separators[0].Location.Curve
        try:
            source_z = first_curve.GetEndPoint(0).Z
        except Exception:
            source_z = target_elevation(target_view, 0.0)
        elevation = target_elevation(target_view, source_z)
        sketch_plane = None
        created = 0
        failed = []
        transaction = DB.Transaction(doc, u"Копировать разделители в границы зон")
        try:
            transaction.Start()
            plane = Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elevation))
            sketch_plane = SketchPlane.Create(doc, plane)
            for separator in separators:
                try:
                    source_curve = separator.Location.Curve
                    curve = copy_curve_to_elevation(source_curve, elevation)
                    if curve is None:
                        failed.append(separator.Id.IntegerValue)
                        continue
                    boundary = create_area_boundary(doc, target_view, curve, sketch_plane)
                    if boundary is None:
                        failed.append(separator.Id.IntegerValue)
                    else:
                        created += 1
                except Exception:
                    failed.append(separator.Id.IntegerValue)
            if created == 0:
                transaction.RollBack()
                forms.alert(u"Не удалось создать ни одной границы зон. Исходные разделители не изменены.",
                            title=u"Разделители в границы зон")
                return
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise

        message = u"Создано границ зон: {0}.\nЦелевой план: {1}.\nСуществующие линии не удалялись.".format(
            created, target_view.Name)
        if failed:
            message += u"\nНе удалось скопировать: {0}.".format(len(failed))
        forms.alert(message, title=u"Разделители в границы зон")
    except Exception as error:
        forms.alert(u"Ошибка копирования границ зон:\n{0}".format(error),
                    title=u"Разделители в границы зон")


if __name__ == u"__main__":
    main()
