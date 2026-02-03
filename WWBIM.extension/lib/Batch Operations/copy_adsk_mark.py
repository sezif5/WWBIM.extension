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
)

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


CONFIG = {
    "PARAMETER_NAME": "ADSK_Марка",
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


def GetMarkParameterValue(element):
    param = element.LookupParameter("ADSK_Марка")
    if param and param.StorageType == StorageType.String and param.HasValue:
        return param.AsString()
    return None


def SetSystemMarkParameter(element, value):
    try:
        param = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
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
    except:
        pass
    return False


def FillMarkParameter(doc, progress_callback=None):
    categories = GetEngineeringCategories()
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

            mark_value = GetMarkParameterValue(element)
            if mark_value is not None:
                all_values.add(mark_value)
                if SetSystemMarkParameter(element, mark_value):
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
        "values": sorted(list(all_values)),
        "filled": filled,
    }


def Execute(doc, progress_callback=None):
    t = Transaction(doc, "Копирование ADSK_Марка в системный параметр Марка")
    t.Start()

    try:
        param_result = EnsureParameterExists(doc)
        fill_result = FillMarkParameter(doc, progress_callback)

        t.Commit()

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": param_result["message"],
            "fill": {
                "target_param": "Марка",
                "source_param": "ADSK_Марка",
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
                "target_param": "Марка",
                "source_param": "ADSK_Марка",
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
