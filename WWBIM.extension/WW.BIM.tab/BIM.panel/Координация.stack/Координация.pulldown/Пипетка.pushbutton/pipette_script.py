# -*- coding: utf-8 -*-
from __future__ import print_function, division

from pyrevit import revit, DB
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException


doc = revit.doc
uidoc = revit.uidoc


def get_excluded_parameter_ids():
    result = set()
    for name in (u"ELEM_FAMILY_AND_TYPE_PARAM", u"ELEM_FAMILY_PARAM", u"ELEM_TYPE_PARAM"):
        try:
            builtin_parameter = getattr(DB.BuiltInParameter, name)
            result.add(DB.ElementId(builtin_parameter).IntegerValue)
        except Exception:
            pass
    return result


EXCLUDED_PARAMETER_IDS = get_excluded_parameter_ids()


class FamilyInstanceFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return isinstance(element, DB.FamilyInstance)
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


class SameFamilyFilter(ISelectionFilter):
    def __init__(self, family_id, source_id):
        self.family_id = family_id
        self.source_id = source_id

    def AllowElement(self, element):
        try:
            if not isinstance(element, DB.FamilyInstance):
                return False
            if element.Id.IntegerValue == self.source_id:
                return False
            return element.Symbol.Family.Id.IntegerValue == self.family_id
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


def get_parameter_id(parameter):
    try:
        return parameter.Id.IntegerValue
    except Exception:
        return None


def get_parameter_key(parameter):
    try:
        return (parameter.Definition.Name, parameter.StorageType)
    except Exception:
        return None


def collect_target_parameters(element):
    by_id = {}
    by_key = {}
    duplicate_keys = set()

    try:
        parameters = element.Parameters
    except Exception:
        parameters = []

    for parameter in parameters:
        parameter_id = get_parameter_id(parameter)
        if parameter_id is not None:
            by_id[parameter_id] = parameter

        key = get_parameter_key(parameter)
        if key is None:
            continue
        if key in by_key:
            duplicate_keys.add(key)
        else:
            by_key[key] = parameter

    for key in duplicate_keys:
        try:
            del by_key[key]
        except Exception:
            pass

    return by_id, by_key


def find_target_parameter(source_parameter, by_id, by_key):
    parameter_id = get_parameter_id(source_parameter)
    if parameter_id is not None and parameter_id in by_id:
        return by_id[parameter_id]

    key = get_parameter_key(source_parameter)
    if key is not None:
        return by_key.get(key)
    return None


def values_equal(source_parameter, target_parameter):
    try:
        storage_type = source_parameter.StorageType
        if storage_type != target_parameter.StorageType:
            return False
        if storage_type == DB.StorageType.String:
            return (source_parameter.AsString() or u"") == (target_parameter.AsString() or u"")
        if storage_type == DB.StorageType.Double:
            return abs(source_parameter.AsDouble() - target_parameter.AsDouble()) < 1e-9
        if storage_type == DB.StorageType.Integer:
            return source_parameter.AsInteger() == target_parameter.AsInteger()
        if storage_type == DB.StorageType.ElementId:
            return source_parameter.AsElementId().IntegerValue == target_parameter.AsElementId().IntegerValue
    except Exception:
        return False
    return False


def set_parameter_value(source_parameter, target_parameter):
    try:
        if target_parameter.IsReadOnly:
            return u"readonly"
        if source_parameter.StorageType != target_parameter.StorageType:
            return u"storage_mismatch"
        if values_equal(source_parameter, target_parameter):
            return u"unchanged"

        storage_type = source_parameter.StorageType
        if storage_type == DB.StorageType.String:
            target_parameter.Set(source_parameter.AsString() or u"")
        elif storage_type == DB.StorageType.Double:
            target_parameter.Set(source_parameter.AsDouble())
        elif storage_type == DB.StorageType.Integer:
            target_parameter.Set(source_parameter.AsInteger())
        elif storage_type == DB.StorageType.ElementId:
            target_parameter.Set(source_parameter.AsElementId())
        else:
            return u"unsupported"
        return u"updated"
    except Exception:
        return u"error"


def copy_instance_parameters(source, target):
    result = {
        u"updated": 0,
        u"unchanged": 0,
        u"skipped": 0,
        u"errors": 0,
    }
    by_id, by_key = collect_target_parameters(target)

    try:
        source_parameters = source.Parameters
    except Exception:
        source_parameters = []

    for source_parameter in source_parameters:
        if get_parameter_id(source_parameter) in EXCLUDED_PARAMETER_IDS:
            result[u"skipped"] += 1
            continue

        target_parameter = find_target_parameter(source_parameter, by_id, by_key)
        if target_parameter is None:
            result[u"skipped"] += 1
            continue

        status = set_parameter_value(source_parameter, target_parameter)
        if status == u"updated":
            result[u"updated"] += 1
        elif status == u"unchanged":
            result[u"unchanged"] += 1
        elif status == u"error":
            result[u"errors"] += 1
        else:
            result[u"skipped"] += 1

    return result


def main():
    try:
        source_reference = uidoc.Selection.PickObject(
            ObjectType.Element,
            FamilyInstanceFilter(),
            u"Пипетка: выберите исходный экземпляр семейства"
        )
    except OperationCanceledException:
        return
    except Exception:
        return

    try:
        source = doc.GetElement(source_reference.ElementId)
        family_id = source.Symbol.Family.Id.IntegerValue
        source_id = source.Id.IntegerValue
    except Exception:
        return

    target_filter = SameFamilyFilter(family_id, source_id)

    while True:
        try:
            target_reference = uidoc.Selection.PickObject(
                ObjectType.Element,
                target_filter,
                u"Пипетка: выбирайте экземпляры того же семейства, Esc — завершить"
            )
            target = doc.GetElement(target_reference.ElementId)
        except OperationCanceledException:
            break
        except Exception:
            break

        transaction = DB.Transaction(doc, u"Пипетка: копировать параметры экземпляра")
        try:
            transaction.Start()
            copy_instance_parameters(source, target)
            transaction_status = transaction.Commit()
            if transaction_status != DB.TransactionStatus.Committed:
                raise Exception(u"Транзакция не была завершена")
        except Exception:
            try:
                if transaction.GetStatus() == DB.TransactionStatus.Started:
                    transaction.RollBack()
            except Exception:
                pass


if __name__ == u"__main__":
    main()
