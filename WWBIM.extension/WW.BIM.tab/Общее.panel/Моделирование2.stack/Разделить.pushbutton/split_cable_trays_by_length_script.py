# -*- coding: utf-8 -*-
from __future__ import division

from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter


doc = revit.doc
uidoc = revit.uidoc


class CableTraySelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            cat = element.Category
            if cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_CableTray):
                return True
        except Exception:
            pass
        return False

    def AllowReference(self, reference, point):
        return True


def mm_to_ft(mm_value):
    # 1 фут = 304.8 мм
    return mm_value / 304.8


def split_cable_tray(tray, segment_length_ft, tol=1.0e-6):
    """Разбить один лоток на отрезки заданной длины.

    Создаёт новые экземпляры CableTray той же марки и уровня,
    копирует базовые параметры, удаляет исходный элемент.
    """
    loc = tray.Location
    if not isinstance(loc, DB.LocationCurve):
        return

    curve = loc.Curve
    if curve is None:
        return

    total_length = curve.Length
    if total_length <= segment_length_ft + tol:
        # Лоток короче или примерно равен требуемой длине — не трогаем
        return

    # Собираем расстояния вдоль кривой, где будем резать
    distances = []
    d = segment_length_ft
    while d < total_length - tol:
        distances.append(d)
        d += segment_length_ft

    if not distances:
        return

    # Точки деления
    points = [curve.GetEndPoint(0)]
    for d in distances:
        param = d / total_length
        pt = curve.Evaluate(param, True)
        points.append(pt)
    points.append(curve.GetEndPoint(1))

    tray_type_id = tray.GetTypeId()

    # Уровень
    level_id = DB.ElementId.InvalidElementId
    try:
        level = tray.ReferenceLevel
        if level:
            level_id = level.Id
    except Exception:
        try:
            par_level = tray.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
            if par_level:
                level_id = par_level.AsElementId()
        except Exception:
            pass

    # Параметры, которые попробуем скопировать
    params_to_copy = [
        DB.BuiltInParameter.RBS_OFFSET_PARAM,
        DB.BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM,
        DB.BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM,
    ]

    new_tray_ids = []

    for i in range(len(points) - 1):
        p0 = points[i]
        p1 = points[i + 1]

        if p0.DistanceTo(p1) < tol:
            continue

        new_tray = DB.Electrical.CableTray.Create(doc, tray_type_id, p0, p1, level_id)
        new_tray_ids.append(new_tray.Id)

        # Копируем выбранные параметры
        for bip in params_to_copy:
            try:
                src_par = tray.get_Parameter(bip)
                dst_par = new_tray.get_Parameter(bip)
                if src_par and dst_par and (not dst_par.IsReadOnly):
                    if src_par.StorageType == DB.StorageType.Double:
                        dst_par.Set(src_par.AsDouble())
                    elif src_par.StorageType == DB.StorageType.Integer:
                        dst_par.Set(src_par.AsInteger())
                    elif src_par.StorageType == DB.StorageType.String:
                        dst_par.Set(src_par.AsString())
                    elif src_par.StorageType == DB.StorageType.ElementId:
                        dst_par.Set(src_par.AsElementId())
            except Exception:
                pass

    # Удаляем исходный лоток, если удалось создать новые
    if new_tray_ids:
        doc.Delete(tray.Id)


def main():
    # Выбор длины сегмента
    options = [
        (u"2000 мм", 2000),
        (u"3000 мм", 3000),
        (u"6000 мм", 6000),
    ]

    labels = [opt[0] for opt in options]

    choice = forms.SelectFromList.show(
        labels,
        title=u"Выберите длину сегментов кабельных лотков",
        multiselect=False
    )

    if not choice:
        return

    selected_mm = None
    for label, mm_value in options:
        if label == choice:
            selected_mm = mm_value
            break

    if selected_mm is None:
        return

    segment_length_ft = mm_to_ft(selected_mm)

    # Выбор лотков пользователем
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            CableTraySelectionFilter(),
            u"Выберите кабельные лотки для деления"
        )
    except Exception:
        return

    trays = [doc.GetElement(r.ElementId) for r in refs]

    if not trays:
        return

    t = DB.Transaction(doc, u"Разделить кабельные лотки")
    t.Start()
    try:
        for tray in trays:
            split_cable_tray(tray, segment_length_ft)
        t.Commit()
    except Exception as ex:
        t.RollBack()
        forms.alert(
            u"Ошибка при делении кабельных лотков:\n{0}".format(ex),
            ok=True
        )


if __name__ == "__main__":
    main()
