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
    Transaction,
    StorageType,
)

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


CONFIG = {
    "PARAMETER_NAME": "ADSK_Имя_файла",
    "BINDING_TYPE": "Instance",
    "PARAMETER_GROUP": "PG_IDENTITY_DATA",
}


def EnsureParameterExists(doc):
    from Autodesk.Revit.DB import BuiltInParameterGroup

    param_config = {
        "PARAMETER_NAME": CONFIG["PARAMETER_NAME"],
        "BINDING_TYPE": CONFIG["BINDING_TYPE"],
        "PARAMETER_GROUP": BuiltInParameterGroup.PG_IDENTITY_DATA,
        "CATEGORIES": MODEL_CATEGORIES,
    }

    try:
        return AddSharedParameterToDoc(doc, param_config)
    except Exception as e:
        return {
            "success": False,
            "parameters": {"added": [], "existing": [], "failed": []},
            "message": "Ошибка при проверке параметра: {0}".format(str(e)),
        }


def GetAllElements(doc):
    collector = FilteredElementCollector(doc)
    return collector.WhereElementIsNotElementType().ToElements()


def GetParameterValue(element, param_name):
    param = element.LookupParameter(param_name)
    if param and param.StorageType == StorageType.String and param.HasValue:
        return param.AsString()
    return None


def SetParameterValue(element, param_name, value):
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return {"status": "parameter_not_found", "reason": "parameter_not_found"}

        if param.StorageType != StorageType.String:
            return {"status": "wrong_storage_type", "reason": "wrong_storage_type"}

        if param.IsReadOnly:
            return {"status": "readonly", "reason": "readonly"}

        try:
            current_value = param.AsString()
            if current_value == value:
                return {"status": "already_ok", "reason": "already_ok"}
        except:
            pass

        if value is not None:
            param.Set(value)
            return {"status": "updated", "reason": None}

        return {"status": "exception", "reason": "value_is_none"}
    except Exception as e:
        return {"status": "exception", "reason": "exception"}


def FillFilenameParameter(doc, progress_callback=None):
    filename = doc.Title
    elements = GetAllElements(doc)
    total = len(elements)
    updated_count = 0
    skipped_count = 0
    skip_reasons = {
        "parameter_not_found": 0,
        "readonly": 0,
        "wrong_storage_type": 0,
        "already_ok": 0,
        "exception": 0,
    }

    if total == 0:
        return {
            "planned_value": filename,
            "total": 0,
            "updated_count": 0,
            "skipped_count": 0,
            "skip_reasons": {},
            "values": [],
            "filled": False,
        }

    current_index = 0
    param_name = CONFIG["PARAMETER_NAME"]

    for element in elements:
        if progress_callback:
            progress = int((current_index / total) * 100)
            progress_callback(progress)

        current_value = GetParameterValue(element, param_name)
        if current_value != filename:
            result = SetParameterValue(element, param_name, filename)
            if result["status"] == "updated":
                updated_count += 1
            else:
                skipped_count += 1
                reason = result["reason"]
                if reason in skip_reasons:
                    skip_reasons[reason] += 1
        else:
            skipped_count += 1

        current_index += 1

    filled = updated_count > 0

    reasons_str = "; ".join(
        ["{0}={1}".format(k, v) for k, v in skip_reasons.items() if v > 0]
    )
    message = "planned={0}, total={1}, updated={2}, skipped={3}".format(
        filename, total, updated_count, skipped_count
    )
    if reasons_str:
        message += "; reasons: " + reasons_str

    return {
        "planned_value": filename,
        "total": total,
        "updated_count": updated_count,
        "skipped_count": skipped_count,
        "skip_reasons": skip_reasons,
        "values": [filename],
        "filled": filled,
        "message": message,
    }


def Execute(doc, progress_callback=None):
    filename = doc.Title
    t = Transaction(doc, "Заполнение параметра ADSK_Имя_файла")
    t.Start()

    try:
        param_result = EnsureParameterExists(doc)
        
        # Проверяем результат добавления параметра
        # Если параметр не добавлен и не существует - это ошибка
        if not param_result.get("success", False):
            # Проверяем, может параметр уже существует (это ok)
            if not param_result.get("parameters", {}).get("existing"):
                t.RollBack()
                return {
                    "success": False,
                    "message": param_result.get("message", "Не удалось добавить параметр"),
                    "parameters": param_result.get("parameters", {"added": [], "existing": [], "failed": []}),
                    "fill": {
                        "target_param": CONFIG["PARAMETER_NAME"],
                        "source": "Название файла: {0}".format(filename),
                        "filled": False,
                        "planned_value": None,
                        "total": 0,
                        "updated_count": 0,
                        "skipped_count": 0,
                        "skip_reasons": {},
                        "values": [],
                    },
                }
        
        fill_result = FillFilenameParameter(doc, progress_callback)

        t.Commit()

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": "Заполнение завершено",
            "fill": {
                "target_param": CONFIG["PARAMETER_NAME"],
                "source": "Название файла: {0}".format(filename),
                "filled": fill_result["filled"],
                "planned_value": fill_result["planned_value"],
                "total": fill_result["total"],
                "updated_count": fill_result["updated_count"],
                "skipped_count": fill_result["skipped_count"],
                "skip_reasons": fill_result["skip_reasons"],
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
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": CONFIG["PARAMETER_NAME"],
                "source": "Название файла: {0}".format(filename),
                "filled": False,
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
