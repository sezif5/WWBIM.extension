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
    "PARAMETER_GROUP": "INVALID",
}


def EnsureParameterExists(doc):
    from Autodesk.Revit.DB import BuiltInParameterGroup

    param_config = {
        "PARAMETER_NAME": CONFIG["PARAMETER_NAME"],
        "BINDING_TYPE": CONFIG["BINDING_TYPE"],
        "PARAMETER_GROUP": BuiltInParameterGroup.INVALID,
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
    param = element.LookupParameter(param_name)
    if param and param.StorageType == StorageType.String:
        try:
            current_value = param.AsString()
            if current_value == value:
                return False
        except:
            pass

        if value is not None:
            param.Set(value)
            return True
    return False


def FillFilenameParameter(doc, progress_callback=None):
    filename = doc.Title
    elements = GetAllElements(doc)
    total_elements = len(elements)
    updated_elements = 0
    skipped_elements = 0

    if total_elements == 0:
        return {
            "total_elements": 0,
            "updated_elements": 0,
            "skipped_elements": 0,
            "values": [],
            "filled": False,
        }

    current_index = 0
    param_name = CONFIG["PARAMETER_NAME"]

    for element in elements:
        if progress_callback:
            progress = int((current_index / total_elements) * 100)
            progress_callback(progress)

        current_value = GetParameterValue(element, param_name)
        if current_value != filename:
            if SetParameterValue(element, param_name, filename):
                updated_elements += 1
            else:
                skipped_elements += 1
        else:
            skipped_elements += 1

        current_index += 1

    filled = updated_elements > 0
    return {
        "total_elements": total_elements,
        "updated_elements": updated_elements,
        "skipped_elements": skipped_elements,
        "values": [filename],
        "filled": filled,
    }


def Execute(doc, progress_callback=None):
    filename = doc.Title
    t = Transaction(doc, "Заполнение параметра ADSK_Имя_файла")
    t.Start()

    try:
        param_result = EnsureParameterExists(doc)
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
                "total_elements": fill_result["total_elements"],
                "updated_elements": fill_result["updated_elements"],
                "skipped_elements": fill_result["skipped_elements"],
                "values": fill_result["values"],
            },
        }

        if not fill_result["filled"] and fill_result["total_elements"] == 0:
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
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
