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
import re

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


CONFIG = {
    "SECTION_PARAMETER": "ADSK_Секция",
    "BINDING_TYPE": "Instance",
    "PARAMETER_GROUP": "PG_IDENTITY_DATA",
    # Regex паттерны для извлечения секции из имени файла
    # Группа (1) должна содержать номер секции
    "PATTERNS": [
        r"_С(\d+)_",           # _С1_, _С02_ и т.д.
        r"_Секция(\d+)_",     # _Секция1_, _Секция02_
        r"_SEC(\d+)_",        # _SEC1_, _SEC02_
        r"-С(\d+)-",          # -С1-, -С02-
    ],
    # Позиция секции в имени файла при разбиении по "_" (0-indexed)
    # Используется если regex паттерны не сработали
    "SECTION_POSITION": 2,
}


def EnsureParameterExists(doc):
    from Autodesk.Revit.DB import BuiltInParameterGroup

    param_config = {
        "PARAMETER_NAME": CONFIG["SECTION_PARAMETER"],
        "BINDING_TYPE": CONFIG["BINDING_TYPE"],
        "PARAMETER_GROUP": BuiltInParameterGroup.PG_IDENTITY_DATA,
        "CATEGORIES": MODEL_CATEGORIES,
    }

    try:
        return AddSharedParameterToDoc(doc, param_config)
    except:
        return {
            "success": True,
            "parameters": {"added": [], "existing": [], "failed": []},
            "message": "Ошибка при проверке параметра",
        }


def ExtractSectionFromFilename(filename):
    for pattern in CONFIG["PATTERNS"]:
        match = re.search(pattern, filename)
        if match:
            return match.group(1)

    parts = filename.split("_")
    parts_count = len(parts)

    section_index = CONFIG["SECTION_POSITION"]
    if section_index < parts_count:
        part = parts[section_index]
        if "-" in part:
            section_part = part.split("-")[-1]
            return section_part
        return part

    return None


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


def FillSectionFromFilename(doc, progress_callback=None):
    filename = doc.Title
    section = ExtractSectionFromFilename(filename)

    if not section:
        return {
            "planned_value": None,
            "total": 0,
            "updated_count": 0,
            "skipped_count": 0,
            "skip_reasons": {},
            "values": [],
            "filled": False,
            "message": "Не удалось определить секцию из названия файла",
        }

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
            "planned_value": section,
            "total": 0,
            "updated_count": 0,
            "skipped_count": 0,
            "skip_reasons": {},
            "values": [],
            "filled": False,
        }

    current_index = 0
    section_param = CONFIG["SECTION_PARAMETER"]

    for element in elements:
        if progress_callback:
            progress = int((current_index / total) * 100)
            progress_callback(progress)

        result = SetParameterValue(element, section_param, section)
        if result["status"] == "updated":
            updated_count += 1
        else:
            skipped_count += 1
            reason = result["reason"]
            if reason in skip_reasons:
                skip_reasons[reason] += 1

        current_index += 1

    filled = updated_count > 0

    reasons_str = "; ".join(
        ["{0}={1}".format(k, v) for k, v in skip_reasons.items() if v > 0]
    )
    message = "planned={0}, total={1}, updated={2}, skipped={3}".format(
        section, total, updated_count, skipped_count
    )
    if reasons_str:
        message += "; reasons: " + reasons_str

    return {
        "planned_value": section,
        "total": total,
        "updated_count": updated_count,
        "skipped_count": skipped_count,
        "skip_reasons": skip_reasons,
        "values": [section],
        "filled": filled,
        "message": message,
    }


def Execute(doc, progress_callback=None):
    filename = doc.Title
    t = Transaction(doc, "Заполнение параметра ADSK_Секция из названия файла")
    t.Start()

    try:
        param_result = EnsureParameterExists(doc)
        fill_result = FillSectionFromFilename(doc, progress_callback)

        if not fill_result["filled"] and "message" in fill_result:
            t.RollBack()
            return {
                "success": False,
                "message": fill_result["message"],
                "parameters": param_result["parameters"],
                "fill": {
                    "target_param": "ADSK_Секция",
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

        t.Commit()

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": param_result["message"],
            "fill": {
                "target_param": "ADSK_Секция",
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
                "target_param": "ADSK_Секция",
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
