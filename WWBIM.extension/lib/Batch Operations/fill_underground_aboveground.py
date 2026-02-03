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
    ViewSheet,
    View,
    Parameter,
    StorageType,
    ElementId,
    ViewType,
    Transaction,
    BoundingBoxXYZ,
)

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


CONFIG = {
    "PARAMETER_NAME": "ADSK_ПодземныйНадземный",
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


def GetAllCategories(doc):
    categories = []
    for category in doc.Settings.Categories:
        if category.CanAddSubcategory:
            categories.append(category)
    return categories


def GetCategoryElements(doc, category):
    try:
        collector = FilteredElementCollector(doc).OfCategory(category.Id.IntegerValue)
        return collector.WhereElementIsNotElementType().ToElements()
    except:
        return []


def GetElementBoundingBox(element):
    try:
        bbox = element.get_BoundingBox(None)
        if not bbox:
            bbox = element.get_BoundingBox(element.Document.ActiveView)
        return bbox
    except:
        return None


def DetermineUndergroundAboveground(element, doc):
    bbox = GetElementBoundingBox(element)
    if not bbox:
        return None

    min_z = bbox.Min.Z

    if min_z >= 0:
        return "Надземный"
    else:
        return "Подземный"


def GetUndergroundAbovegroundParameter(element):
    param = element.LookupParameter("ADSK_ПодземныйНадземный")
    if param and param.StorageType == StorageType.String:
        try:
            return param.AsString()
        except:
            return None
    return None


def SetUndergroundAbovegroundParameter(element, value):
    param = element.LookupParameter("ADSK_ПодземныйНадземный")
    if param and param.StorageType == StorageType.String:
        try:
            current_value = param.AsString()
            if current_value == value:
                return False
        except:
            pass

        if value:
            param.Set(value)
            return True
    return False


def FillUndergroundAboveground(doc, progress_callback=None):
    categories = GetAllCategories(doc)
    total_elements = 0
    updated_elements = 0
    skipped_elements = 0
    all_values = set()

    for category in categories:
        elements = GetCategoryElements(doc, category)
        total_elements += len(elements)

    if total_elements == 0:
        return {
            "total_elements": 0,
            "updated_elements": 0,
            "skipped_elements": 0,
            "values": [],
            "filled": False,
        }

    current_index = 0

    for category in categories:
        elements = GetCategoryElements(doc, category)

        for element in elements:
            if progress_callback:
                progress = int((current_index / total_elements) * 100)
                progress_callback(progress)

            value = DetermineUndergroundAboveground(element, doc)
            if value:
                all_values.add(value)
                current_value = GetUndergroundAbovegroundParameter(element)
                if current_value != value:
                    if SetUndergroundAbovegroundParameter(element, value):
                        updated_elements += 1
                    else:
                        skipped_elements += 1
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
        "values": sorted(list(all_values)),
        "filled": filled,
    }


def Execute(doc, progress_callback=None):
    t = Transaction(doc, "Заполнение параметра ADSK_ПодземныйНадземный")
    t.Start()

    try:
        param_result = EnsureParameterExists(doc)
        fill_result = FillUndergroundAboveground(doc, progress_callback)

        t.Commit()

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": param_result["message"],
            "fill": {
                "target_param": "ADSK_ПодземныйНадземный",
                "source": "Координата Z элемента",
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
                "target_param": "ADSK_ПодземныйНадземный",
                "source": "Координата Z элемента",
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
