# -*- coding: utf-8 -*-
from __future__ import print_function

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB.ExtensibleStorage import AccessLevel, Entity, Schema, SchemaBuilder
from System import Guid, String
from System.Collections.Generic import List


SCHEMA_ID = Guid("8706d765-839a-4307-92c8-4c9348f5c086")
FIELD_NAME = "SourceAreaUniqueId"


def get_schema(create=False):
    schema = Schema.Lookup(SCHEMA_ID)
    if schema is not None or not create:
        return schema
    builder = SchemaBuilder(SCHEMA_ID)
    builder.SetSchemaName("WWBIM_AreaHatches_v1")
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.AddSimpleField(FIELD_NAME, String)
    return builder.Finish()


def find_solid_type(doc):
    for region_type in DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType):
        try:
            pattern = doc.GetElement(region_type.ForegroundPatternId)
            if pattern is not None and pattern.GetFillPattern().IsSolidFill:
                return region_type
        except Exception:
            continue
    return None


def get_area_loops(area):
    options = DB.SpatialElementBoundaryOptions()
    segments = area.GetBoundarySegments(options)
    if not segments:
        raise ValueError(u"У площади нет расчетного контура.")
    loops = List[DB.CurveLoop]()
    for contour in segments:
        loop = DB.CurveLoop()
        count = 0
        for segment in contour:
            curve = segment.GetCurve()
            if curve is None or not curve.IsBound:
                raise ValueError(u"Граница площади содержит незамкнутый сегмент.")
            loop.Append(curve)
            count += 1
        if count:
            if loop.IsOpen():
                raise ValueError(u"Граница площади не замкнута.")
            loops.Add(loop)
    if loops.Count == 0:
        raise ValueError(u"У площади нет замкнутых расчетных контуров.")
    return loops


def area_label(area):
    try:
        return u"{0} (ID {1})".format(area.Name, area.Id.IntegerValue)
    except Exception:
        return u"ID {0}".format(area.Id.IntegerValue)


def existing_regions(doc, view):
    schema = get_schema()
    if schema is None:
        return []
    field = schema.GetField(FIELD_NAME)
    result = []
    for region in DB.FilteredElementCollector(doc, view.Id).OfClass(DB.FilledRegion):
        try:
            entity = region.GetEntity(schema)
            if entity.IsValid() and entity.Get[String](field):
                result.append(region.Id)
        except Exception:
            continue
    return result


def main():
    doc = revit.doc
    view = revit.active_view
    try:
        if doc is None or view is None or view.ViewType != DB.ViewType.AreaPlan or view.IsTemplate:
            forms.alert(u"Откройте план площадей (план зонирования) и повторите команду.", title=u"Штриховка зон")
            return
        if doc.IsReadOnly:
            forms.alert(u"Документ доступен только для чтения.", title=u"Штриховка зон")
            return
        region_type = find_solid_type(doc)
        if region_type is None:
            forms.alert(u"В проекте нет типа закрашенной области со сплошным узором переднего плана. Создайте его в Revit и повторите команду.",
                        title=u"Штриховка зон")
            return

        areas = []
        for element in DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_Areas).WhereElementIsNotElementType():
            try:
                if isinstance(element, DB.Area) and element.Area > 0:
                    areas.append(element)
            except Exception:
                continue
        if not areas:
            forms.alert(u"На активном плане нет размещённых площадей с рассчитанной площадью. Старые штриховки не удалены.",
                        title=u"Штриховка зон")
            return

        boundaries = []
        for area in areas:
            try:
                boundaries.append((area.UniqueId, get_area_loops(area)))
            except Exception as error:
                raise ValueError(u"Не удалось получить контур зоны {0}: {1}".format(area_label(area), error))

        old_ids = existing_regions(doc, view)
        transaction = DB.Transaction(doc, u"Обновить штриховку зон")
        try:
            transaction.Start()
            schema = get_schema(True)
            field = schema.GetField(FIELD_NAME)
            if old_ids:
                ids = List[DB.ElementId]()
                for old_id in old_ids:
                    ids.Add(old_id)
                deleted = doc.Delete(ids)
                allowed = set(old_id.IntegerValue for old_id in old_ids)
                if any(item.IntegerValue not in allowed for item in deleted):
                    raise RuntimeError(u"От прежних штриховок зависят другие элементы; обновление отменено.")
            for source_id, loops in boundaries:
                region = DB.FilledRegion.Create(doc, region_type.Id, view.Id, loops)
                entity = Entity(schema)
                entity.Set[String](field, source_id)
                region.SetEntity(entity)
            if transaction.Commit() != DB.TransactionStatus.Committed:
                raise RuntimeError(u"Revit не подтвердил изменения.")
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        forms.alert(u"Создано штриховок: {0}. Обновлено прежних: {1}. Изменения сделаны только на активном плане.".format(
            len(boundaries), len(old_ids)), title=u"Штриховка зон")
    except Exception as error:
        forms.alert(u"Не удалось обновить штриховки. Прежние штриховки сохранены.\n{0}".format(error),
                    title=u"Штриховка зон")


if __name__ == u"__main__":
    main()
