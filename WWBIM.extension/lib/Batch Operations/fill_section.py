# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, ".."))
sys.path.insert(0, LIB_DIR)

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInParameter,
    Transaction,
    StorageType,
)

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


# ---------- Пути ----------


def _norm(p):
    return os.path.normpath(os.path.abspath(p)) if p else p


def _module_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def _find_scripts_root(start_dir):
    cur = _norm(start_dir)
    for _ in range(0, 8):
        if os.path.basename(cur).lower() == "scripts":
            return cur
        parent = os.path.dirname(cur)
        if not parent or parent == cur:
            break
        cur = parent
    return _norm(os.path.join(start_dir, os.pardir, os.pardir))


SCRIPTS_ROOT = _find_scripts_root(_module_dir())
OBJECTS_DIR = _norm(os.path.join(SCRIPTS_ROOT, "Objects"))


CONFIG = {
    "ALBUM_PARAMETER": "ADSK_КомплектШифр",
    "SECTION_PARAMETER": "ADSK_Секция",
    "MAPPING_FILE": None,
    "PARAMETER_NAME": "ADSK_Секция",
    "BINDING_TYPE": "Instance",
    "PARAMETER_GROUP": "PG_IDENTITY_DATA",
    "CATEGORIES": MODEL_CATEGORIES,
}

# Заполняем MAPPING_FILE абсолютным путём
CONFIG["MAPPING_FILE"] = os.path.join(OBJECTS_DIR, "album_section_mapping.txt")


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
        with open(mapping_file_path, "r", encoding="utf-8") as f:
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


def FillSectionParameter(doc, mapping, progress_callback=None):
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
    all_values = set()

    if total == 0:
        return {
            "planned_value": None,
            "total": 0,
            "updated_count": 0,
            "skipped_count": 0,
            "skip_reasons": {},
            "values": [],
            "filled": False,
        }

    current_index = 0
    album_param = CONFIG["ALBUM_PARAMETER"]
    section_param = CONFIG["SECTION_PARAMETER"]

    for element in elements:
        if progress_callback:
            progress = int((current_index / total) * 100)
            progress_callback(progress)

        album_value = GetParameterValue(element, album_param)
        if album_value and album_value in mapping:
            section_value = mapping[album_value]
            all_values.add(section_value)
            result = SetParameterValue(element, section_param, section_value)
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
        sorted(list(all_values)), total, updated_count, skipped_count
    )
    if reasons_str:
        message += "; reasons: " + reasons_str

    return {
        "planned_value": sorted(list(all_values)),
        "total": total,
        "updated_count": updated_count,
        "skipped_count": skipped_count,
        "skip_reasons": skip_reasons,
        "values": sorted(list(all_values)),
        "filled": filled,
        "message": message,
    }


def Execute(doc, progress_callback=None):
    mapping_file = CONFIG["MAPPING_FILE"]

    if not os.path.exists(mapping_file):
        return {
            "success": False,
            "message": "Файл соответствия не найден: {0}".format(mapping_file),
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": "ADSK_Секция",
                "source": CONFIG["MAPPING_FILE"],
                "filled": False,
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
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
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
                "message": "Файл соответствия пуст",
            },
        }

    t = Transaction(doc, "Заполнение параметра ADSK_Секция")
    t.Start()

    try:
        param_result = EnsureParameterExists(doc)
        
        # Проверяем результат добавления параметра
        if not param_result.get("success", False):
            if not param_result.get("parameters", {}).get("existing"):
                t.RollBack()
                return {
                    "success": False,
                    "message": param_result.get("message", "Не удалось добавить параметр"),
                    "parameters": param_result.get("parameters", {"added": [], "existing": [], "failed": []}),
                    "fill": {
                        "target_param": "ADSK_Секция",
                        "source": CONFIG["MAPPING_FILE"],
                        "filled": False,
                        "planned_value": None,
                        "total": 0,
                        "updated_count": 0,
                        "skipped_count": 0,
                        "skip_reasons": {},
                        "values": [],
                        "mapping_count": len(mapping),
                        "message": "Не удалось добавить параметр",
                    },
                }
        
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
                "planned_value": fill_result["planned_value"],
                "total": fill_result["total"],
                "updated_count": fill_result["updated_count"],
                "skipped_count": fill_result["skipped_count"],
                "skip_reasons": fill_result["skip_reasons"],
                "values": fill_result["values"],
                "mapping_count": len(mapping),
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
                "source": CONFIG["MAPPING_FILE"],
                "filled": False,
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
                "mapping_count": len(mapping),
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
