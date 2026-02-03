# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import os
import inspect

try:
    script_path = inspect.getfile(inspect.currentframe())
    lib_dir = os.path.dirname(os.path.dirname(script_path))
    SCRIPT_DIR = os.path.dirname(script_path)
except:
    lib_dir = os.path.dirname(os.getcwd())
    SCRIPT_DIR = os.getcwd()

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
    "ALBUM_PARAMETER": "ADSK_КомплектШифр",
    "SECTION_PARAMETER": "ADSK_Секция",
    "MAPPING_FILE": "..\\Objects\\album_section_mapping.txt",
    "PARAMETER_NAME": "ADSK_Секция",
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
    except:
        return {
            "success": True,
            "parameters": {"added": [], "existing": [], "failed": []},
            "message": "Ошибка при проверке параметра",
        }


def ReadMappingFile(mapping_file_path):
    mapping = {}

    if not os.path.exists(mapping_file_path):
        return mapping

    try:
        with open(mapping_file_path, "r", encoding="tf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue

                if "-" in line:
                    parts = line.split("-", 1)
                    if len(parts) == 2:
                        album = parts[0].strip()
                        section = parts[1].strip()
                        if album and section:
                            mapping[album] = section
    except Exception as e:
        pass

    return mapping


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


def FillSectionParameter(doc, mapping, progress_callback=None):
    elements = GetAllElements(doc)
    total_elements = len(elements)
    updated_elements = 0
    skipped_elements = 0
    all_values = set()

    if total_elements == 0:
        return {
            "total_elements": 0,
            "updated_elements": 0,
            "skipped_elements": 0,
            "values": [],
            "filled": False,
        }

    current_index = 0
    album_param = CONFIG["ALBUM_PARAMETER"]
    section_param = CONFIG["SECTION_PARAMETER"]

    for element in elements:
        if progress_callback:
            progress = int((current_index / total_elements) * 100)
            progress_callback(progress)

        album_value = GetParameterValue(element, album_param)
        if album_value and album_value in mapping:
            section_value = mapping[album_value]
            all_values.add(section_value)
            if SetParameterValue(element, section_param, section_value):
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
    mapping_file = os.path.join(SCRIPT_DIR, CONFIG["MAPPING_FILE"])

    if not os.path.exists(mapping_file):
        return {
            "success": False,
            "message": "Файл соответствия не найден: {0}".format(mapping_file),
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": "ADSK_Секция",
                "source": CONFIG["MAPPING_FILE"],
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
                "message": "Файл соответствия не найден",
            },
        }

    mapping = ReadMappingFile(mapping_file)
    if not mapping:
        return {
            "success": False,
            "message": "Файл соответствия пуст или содержит ошибки",
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": "ADSK_Секция",
                "source": CONFIG["MAPPING_FILE"],
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
                "message": "Файл соответствия пуст",
            },
        }

    t = Transaction(doc, "Заполнение параметра ADSK_Секция")
    t.Start()

    try:
        param_result = EnsureParameterExists(doc)
        fill_result = FillSectionParameter(doc, mapping, progress_callback)

        t.Commit()

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": param_result["message"],
            "fill": {
                "target_param": "ADSK_Секция",
                "source": CONFIG["MAPPING_FILE"],
                "filled": fill_result["filled"],
                "total_elements": fill_result["total_elements"],
                "updated_elements": fill_result["updated_elements"],
                "skipped_elements": fill_result["skipped_elements"],
                "values": fill_result["values"],
                "mapping_count": len(mapping),
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
                "target_param": "ADSK_Секция",
                "source": CONFIG["MAPPING_FILE"],
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
                "mapping_count": len(mapping),
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
