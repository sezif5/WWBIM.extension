# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import os
import inspect
import re

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
    Level,
    Options,
    ViewDetailLevel,
    ElementMulticategoryFilter,
    Solid,
    GeometryInstance,
    Curve,
    XYZ,
    BoundingBoxXYZ,
    LocationPoint,
    LocationCurve,
    FamilyInstance,
)
from System.Collections.Generic import List

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


CONFIG = {
    "PARAMETER_NAME": "ADSK_Этаж",
    "BINDING_TYPE": "Instance",
    "PARAMETER_GROUP": "PG_IDENTITY_DATA",
    "DEFAULT_OFFSET_MM": 100,
}

PROBLEMATIC_SYMBOLS_CACHE = set()


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


def GetElementsToProcess(doc):
    collector = FilteredElementCollector(doc)
    cats = List[BuiltInCategory]()
    for bic in MODEL_CATEGORIES:
        cats.Add(bic)
    category_filter = ElementMulticategoryFilter(cats)
    return collector.WhereElementIsNotElementType().WherePasses(category_filter)


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


def IsImportInFamily(element):
    try:
        if not element.Category:
            return False

        category_id = element.Category.Id.IntegerValue
        return category_id == int(BuiltInCategory.OST_ImportObjectStyles)
    except Exception:
        return False


def HasImportedCADGeometry(fam_inst):
    try:
        if not isinstance(fam_inst, FamilyInstance):
            return False

        symbol = fam_inst.Symbol
        if not symbol:
            return False

        symbol_id = symbol.Id.IntegerValue
        if symbol_id in PROBLEMATIC_SYMBOLS_CACHE:
            return True

        family = symbol.Family
        if not family:
            return False

        family_name = family.Name.upper() if family.Name else ""
        symbol_name = symbol.Name.upper() if symbol.Name else ""

        cad_keywords = ["DWG", "DXF", "CAD", "IMPORT", "ИМПОРТ"]

        for keyword in cad_keywords:
            if keyword in family_name or keyword in symbol_name:
                PROBLEMATIC_SYMBOLS_CACHE.add(symbol_id)
                return True

        return False
    except Exception:
        return False


def GetGeometryMinZ(element):
    try:
        opt = Options()
        opt.DetailLevel = ViewDetailLevel.Coarse
        opt.IncludeNonVisibleObjects = False

        geo = element.get_Geometry(opt)
        if not geo:
            return None

        def process_geometry(geometry):
            min_z = None

            for geom_obj in geometry:
                if isinstance(geom_obj, Solid):
                    if geom_obj.Volume > 1e-9:
                        try:
                            if min_z is None or geom_obj.BoundingBox[0].Z < min_z:
                                min_z = geom_obj.BoundingBox[0].Z
                        except Exception:
                            pass
                elif isinstance(geom_obj, GeometryInstance):
                    trans_geom = geom_obj.GetInstanceGeometry()
                    if trans_geom:
                        sub_min = process_geometry(trans_geom)
                        if sub_min is not None:
                            if min_z is None or sub_min < min_z:
                                min_z = sub_min
                elif hasattr(geom_obj, "GetEnumerator"):
                    sub_min = process_geometry(geom_obj)
                    if sub_min is not None:
                        if min_z is None or sub_min < min_z:
                            min_z = sub_min

            return min_z

        return process_geometry(geo)
    except Exception:
        return None


def GetLocationElevation(element):
    try:
        location = element.Location
        if not location:
            return None

        if isinstance(location, LocationPoint):
            return location.Point.Z

        if isinstance(location, LocationCurve):
            curve = location.Curve
            if curve:
                return min(curve.GetEndPoint(0).Z, curve.GetEndPoint(1).Z)

        return None
    except Exception:
        return None


def GetElementElevation(element):
    try:
        if IsImportInFamily(element):
            loc_z = GetLocationElevation(element)
            if loc_z is not None:
                return loc_z

            bbox = element.get_BoundingBox(None)
            if bbox:
                return bbox.Min.Z

            return None

        loc_z = GetLocationElevation(element)
        if loc_z is not None:
            return loc_z

        bbox = element.get_BoundingBox(None)
        if bbox:
            return bbox.Min.Z

        if isinstance(element, FamilyInstance) and not HasImportedCADGeometry(element):
            min_z = GetGeometryMinZ(element)
            if min_z is not None:
                return min_z

        return None
    except Exception:
        return None


def ParseLevelName(level_name):
    if level_name == "Уровень 0":
        return None

    pattern = r"^SIM\d+_\d+_AR_(-?\d+)"
    match = re.search(pattern, level_name)

    if match:
        floor_num = match.group(1)
        return str(int(floor_num))

    pattern = r"^KR_(-?\d+)_"
    match = re.search(pattern, level_name)

    if match:
        floor_num = match.group(1)
        return str(int(floor_num))

    return level_name


def GetLevelsOrdered(doc):
    try:
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        return sorted(levels, key=lambda l: l.Elevation)
    except Exception:
        return []


def DetermineFloorForElement(element, levels_sorted, offset_mm=100):
    z_abs = GetElementElevation(element)

    if z_abs is None:
        return None, "no_elevation"

    offset_feet = offset_mm / 304.8
    z_for_level = z_abs + offset_feet

    levels = levels_sorted

    if not levels:
        return None, "no_levels"

    levels_below = [l for l in levels if l.Elevation <= z_for_level]

    if levels_below:
        closest_level = max(levels_below, key=lambda l: l.Elevation)
    else:
        closest_level = min(levels, key=lambda l: abs(l.Elevation - z_for_level))

    floor_value = ParseLevelName(closest_level.Name)

    return floor_value, None


def FillFloorParameter(doc, progress_callback=None):
    global PROBLEMATIC_SYMBOLS_CACHE
    PROBLEMATIC_SYMBOLS_CACHE.clear()

    levels_sorted = GetLevelsOrdered(doc)

    elements = GetElementsToProcess(doc)
    total = elements.GetElementCount()
    updated_count = 0
    skipped_count = 0
    skip_reasons = {
        "parameter_not_found": 0,
        "readonly": 0,
        "wrong_storage_type": 0,
        "already_ok": 0,
        "no_levels": 0,
        "no_elevation": 0,
        "exception": 0,
        "cad_import": 0,
    }
    all_values = set()
    offset_mm = CONFIG["DEFAULT_OFFSET_MM"]

    if total == 0:
        return {
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
            progress = int(((current_index + 1) / float(total)) * 100)
            progress_callback(progress)

        if IsImportInFamily(element):
            skipped_count += 1
            if "cad_import" in skip_reasons:
                skip_reasons["cad_import"] += 1
            else:
                skip_reasons["cad_import"] = 1
            current_index += 1
            continue

        floor_value, skip_reason = DetermineFloorForElement(
            element, levels_sorted, offset_mm
        )

        if floor_value:
            all_values.add(floor_value)
            result = SetParameterValue(element, param_name, floor_value)
            if result["status"] == "updated":
                updated_count += 1
            else:
                skipped_count += 1
                reason = result["reason"]
                if reason in skip_reasons:
                    skip_reasons[reason] += 1
        else:
            skipped_count += 1
            if skip_reason in skip_reasons:
                skip_reasons[skip_reason] += 1
            else:
                skip_reasons["exception"] += 1

        current_index += 1

    filled = updated_count > 0

    reasons_str = "; ".join(
        ["{0}={1}".format(k, v) for k, v in skip_reasons.items() if v > 0]
    )
    message = "total={0}, updated={1}, skipped={2}".format(
        total, updated_count, skipped_count
    )
    if reasons_str:
        message += "; reasons: " + reasons_str

    return {
        "total": total,
        "updated_count": updated_count,
        "skipped_count": skipped_count,
        "skip_reasons": skip_reasons,
        "values": sorted(list(all_values)),
        "filled": filled,
        "message": message,
    }


def Execute(doc, progress_callback=None):
    t = Transaction(doc, "Заполнение параметра ADSK_Этаж")
    t.Start()

    try:
        param_result = EnsureParameterExists(doc)

        if not param_result.get("success", False):
            if not param_result.get("parameters", {}).get("existing"):
                t.RollBack()
                return {
                    "success": False,
                    "message": param_result.get(
                        "message", "Не удалось добавить параметр"
                    ),
                    "parameters": param_result.get(
                        "parameters", {"added": [], "existing": [], "failed": []}
                    ),
                    "fill": {
                        "target_param": CONFIG["PARAMETER_NAME"],
                        "filled": False,
                        "total": 0,
                        "updated_count": 0,
                        "skipped_count": 0,
                        "skip_reasons": {},
                        "values": [],
                    },
                }

        fill_result = FillFloorParameter(doc, progress_callback)

        t.Commit()

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": "Заполнение завершено",
            "fill": {
                "target_param": CONFIG["PARAMETER_NAME"],
                "filled": fill_result["filled"],
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
            skip_reasons = fill_result["skip_reasons"]
            has_other_skip_reasons = any(
                skip_reasons.get(key, 0) > 0
                for key in [
                    "parameter_not_found",
                    "readonly",
                    "wrong_storage_type",
                    "no_levels",
                    "no_elevation",
                    "exception",
                ]
            )
            if not has_other_skip_reasons:
                result["fill"]["message"] = (
                    "Заполнение не требовалось: все элементы уже имели правильные значения"
                )
            else:
                reasons = "; ".join(
                    ["{0}={1}".format(k, v) for k, v in skip_reasons.items() if v > 0]
                )
                result["fill"]["message"] = (
                    "Обновлений нет: элементы пропущены ({0})".format(reasons)
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
                "filled": False,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
