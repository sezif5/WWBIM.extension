# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import os
import inspect

try:
    script_path = inspect.getfile(inspect.currentframe())
    lib_dir = os.path.dirname(os.path.dirname(script_path))
except:
    lib_dir = os.path.dirname(os.getcwd())

if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInParameter,
    BuiltInCategory,
    Transaction,
    StorageType,
    ElementId,
)


def GetEngineeringCategories():
    engineering_categories = [
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PlumbingFixtures,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_CommunicationDevices,
        BuiltInCategory.OST_DuctTerminal,
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_PipeAccessory,
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_Sprinklers,
        BuiltInCategory.OST_FireAlarmDevices,
        BuiltInCategory.OST_CableTray,
        BuiltInCategory.OST_Conduit,
        BuiltInCategory.OST_Wire,
        BuiltInCategory.OST_CableTrayFitting,
        BuiltInCategory.OST_ConduitFitting,
    ]
    return engineering_categories


def GetCategoryElements(doc, category):
    try:
        collector = FilteredElementCollector(doc).OfCategory(category)
        return collector.WhereElementIsNotElementType().ToElements()
    except:
        return []


def GetMarkParameterValue(element, doc):
    param = element.LookupParameter("ADSK_Марка")
    if param and param.StorageType == StorageType.String and param.HasValue:
        value = param.AsString()
        if value and value.strip():
            return value, "instance"

    type_id = element.GetTypeId()
    if type_id and type_id != ElementId.InvalidElementId:
        type_elem = doc.GetElement(type_id)
        if type_elem:
            param = type_elem.LookupParameter("ADSK_Марка")
            if param and param.StorageType == StorageType.String and param.HasValue:
                value = param.AsString()
                if value and value.strip():
                    return value, "type"

    return None, None


def SetSystemMarkParameter(element, value):
    try:
        param = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if not param:
            return {"status": "skipped", "reason": "target_missing"}

        if param.IsReadOnly:
            return {"status": "skipped", "reason": "target_readonly"}

        try:
            current_value = param.AsString()
            if current_value == value:
                return {"status": "skipped", "reason": "already_ok"}
        except:
            pass

        if value is not None:
            param.Set(value)
            return {"status": "updated", "reason": None}

        return {"status": "skipped", "reason": "value_is_none"}
    except Exception as e:
        return {"status": "skipped", "reason": "exception"}


def FillMarkParameter(doc, progress_callback=None):
    categories = GetEngineeringCategories()
    total = 0
    updated_count = 0
    skipped_count = 0
    skip_reasons = {
        "source_missing": 0,
        "source_wrong_type": 0,
        "target_missing": 0,
        "target_readonly": 0,
        "already_ok": 0,
        "exception": 0,
    }
    source_from_instance_count = 0
    source_from_type_count = 0
    all_values = set()

    for category in categories:
        elements = GetCategoryElements(doc, category)
        total += len(elements)

    if total == 0:
        return {
            "planned_value": None,
            "total": 0,
            "updated_count": 0,
            "skipped_count": 0,
            "skip_reasons": {},
            "source_from_instance_count": 0,
            "source_from_type_count": 0,
            "values": [],
            "filled": False,
        }

    current_index = 0

    for category in categories:
        elements = GetCategoryElements(doc, category)

        for element in elements:
            if progress_callback:
                progress = int((current_index / total) * 100)
                progress_callback(progress)

            mark_value, source_type = GetMarkParameterValue(element, doc)
            if mark_value is not None:
                all_values.add(mark_value)
                if source_type == "instance":
                    source_from_instance_count += 1
                elif source_type == "type":
                    source_from_type_count += 1

                result = SetSystemMarkParameter(element, mark_value)
                if result["status"] == "updated":
                    updated_count += 1
                else:
                    skipped_count += 1
                    reason = result["reason"]
                    if reason in skip_reasons:
                        skip_reasons[reason] += 1
            else:
                skipped_count += 1
                skip_reasons["source_missing"] += 1

            current_index += 1

    filled = updated_count > 0

    reasons_str = "; ".join(
        ["{0}={1}".format(k, v) for k, v in skip_reasons.items() if v > 0]
    )
    message = "planned={0}, total={1}, updated={2}, skipped={3}".format(
        sorted(list(all_values)), total, updated_count, skipped_count
    )
    if reasons_str:
        message += "; reasons: " + reasons_str
    message += "; source: instance={0}, type={1}".format(
        source_from_instance_count, source_from_type_count
    )

    return {
        "planned_value": sorted(list(all_values)),
        "total": total,
        "updated_count": updated_count,
        "skipped_count": skipped_count,
        "skip_reasons": skip_reasons,
        "source_from_instance_count": source_from_instance_count,
        "source_from_type_count": source_from_type_count,
        "values": sorted(list(all_values)),
        "filled": filled,
        "message": message,
    }


def Execute(doc, progress_callback=None):
    t = Transaction(doc, "Копирование ADSK_Марка в системный параметр Марка")
    t.Start()

    try:
        fill_result = FillMarkParameter(doc, progress_callback)

        t.Commit()

        result = {
            "success": True,
            "fill": {
                "target_param": "Марка",
                "source_param": "ADSK_Марка",
                "filled": fill_result["filled"],
                "planned_value": fill_result["planned_value"],
                "total": fill_result["total"],
                "updated_count": fill_result["updated_count"],
                "skipped_count": fill_result["skipped_count"],
                "skip_reasons": fill_result["skip_reasons"],
                "source_from_instance_count": fill_result["source_from_instance_count"],
                "source_from_type_count": fill_result["source_from_type_count"],
                "values": fill_result["values"],
                "message": fill_result["message"],
            },
        }

        if not fill_result["filled"] and fill_result["total"] == 0:
            result["fill"]["message"] = (
                "Заполнение не требовалось: нет элементов для обработки"
            )
        elif not fill_result["filled"]:
            result["fill"]["message"] = (
                "Заполнение не требовалось: все элементы уже имели правильные значения"
            )

        return result

    except Exception as e:
        t.RollBack()
        return {
            "success": False,
            "message": "Ошибка: {0}".format(str(e)),
            "fill": {
                "target_param": "Марка",
                "source_param": "ADSK_Марка",
                "filled": False,
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "source_from_instance_count": 0,
                "source_from_type_count": 0,
                "values": [],
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
