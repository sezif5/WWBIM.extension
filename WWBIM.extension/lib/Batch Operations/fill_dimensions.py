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
    ElementMulticategoryFilter,
)
from System.Collections.Generic import List

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


DIMENSION_PARAMETERS = [
    {
        "NAME": "ADSK_Размер_Объём",
        "UNIT_TYPE": "volume",
        "SOURCES": [
            {
                "BIP": BuiltInParameter.HOST_VOLUME_COMPUTED,
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_Ceilings,
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": BuiltInParameter.STRUCTURAL_VOLUME,
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_PIPE_VOLUME_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_DUCT_VOLUME_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                ],
            },
        ],
    },
    {
        "NAME": "ADSK_Размер_Длина",
        "UNIT_TYPE": "length",
        "SOURCES": [
            {
                "BIP": BuiltInParameter.CURVE_ELEM_LENGTH,
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Ramps,
                    BuiltInCategory.OST_CurtainWallMullions,
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
            {
                "BIP": BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH,
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
            {
                "BIP": BuiltInParameter.INSTANCE_LENGTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_PIPE_LENGTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_DUCT_LENGTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_CABLETRAY_LENGTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_CableTray,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_CONDUIT_LENGTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_Conduit,
                ],
            },
        ],
    },
    {
        "NAME": "ADSK_Размер_Ширина",
        "UNIT_TYPE": "length",
        "SOURCES": [
            {
                "BIP": BuiltInParameter.WALL_USER_WIDTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                ],
            },
            {
                "BIP": BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_Floors,
                ],
            },
            {
                "BIP": BuiltInParameter.CEILING_THICKNESS,
                "CATEGORIES": [
                    BuiltInCategory.OST_Ceilings,
                ],
            },
            {
                "BIP": BuiltInParameter.STRUCTURAL_FOUNDATION_THICKNESS,
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": BuiltInParameter.FAMILY_WIDTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_Doors,
                    BuiltInCategory.OST_Windows,
                    BuiltInCategory.OST_Furniture,
                    BuiltInCategory.OST_FurnitureSystems,
                    BuiltInCategory.OST_Casework,
                    BuiltInCategory.OST_SpecialityEquipment,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_CurtainWallPanels,
                    BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_PlumbingFixtures,
                    BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_DUCT_WIDTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctAccessory,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_CableTray,
                ],
            },
        ],
    },
    {
        "NAME": "ADSK_Размер_Высота",
        "UNIT_TYPE": "length",
        "SOURCES": [
            {
                "BIP": BuiltInParameter.FAMILY_HEIGHT_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_Doors,
                    BuiltInCategory.OST_Windows,
                    BuiltInCategory.OST_Furniture,
                    BuiltInCategory.OST_FurnitureSystems,
                    BuiltInCategory.OST_Casework,
                    BuiltInCategory.OST_SpecialityEquipment,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_CurtainWallPanels,
                    BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_PlumbingFixtures,
                    BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_DUCT_HEIGHT_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctAccessory,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_CableTray,
                ],
            },
        ],
    },
    {
        "NAME": "ADSK_Размер_Толщина",
        "UNIT_TYPE": "length",
        "SOURCES": [
            {
                "BIP": BuiltInParameter.WALL_USER_WIDTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                ],
            },
            {
                "BIP": BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_Floors,
                ],
            },
            {
                "BIP": BuiltInParameter.CEILING_THICKNESS,
                "CATEGORIES": [
                    BuiltInCategory.OST_Ceilings,
                ],
            },
            {
                "BIP": BuiltInParameter.STRUCTURAL_FOUNDATION_THICKNESS,
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": BuiltInParameter.FAMILY_DEPTH_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_Doors,
                    BuiltInCategory.OST_Windows,
                    BuiltInCategory.OST_Furniture,
                    BuiltInCategory.OST_Casework,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
        ],
    },
    {
        "NAME": "ADSK_Площадь",
        "UNIT_TYPE": "area",
        "SOURCES": [
            {
                "BIP": BuiltInParameter.HOST_AREA_COMPUTED,
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_Ceilings,
                ],
            },
            {
                "BIP": BuiltInParameter.STRUCTURAL_AREA,
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
        ],
    },
    {
        "NAME": "ADSK_Размер_Диаметр",
        "UNIT_TYPE": "length",
        "SOURCES": [
            {
                "BIP": BuiltInParameter.RBS_PIPE_DIAMETER_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_PIPE_OUTER_DIAMETER_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory,
                ],
            },
            {
                "BIP": BuiltInParameter.RBS_DUCT_DIAMETER_PARAM,
                "CATEGORIES": [
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctAccessory,
                ],
            },
        ],
    },
]


FEET_TO_MM = 304.8
FEET2_TO_M2 = 0.092903
FEET3_TO_M3 = 0.0283168

GENERAL_CONSTRUCTION_CATEGORIES = {
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
}

GENERAL_DIMENSION_PARAMETERS = {
    "ADSK_Размер_Длина",
    "ADSK_Размер_Ширина",
    "ADSK_Размер_Высота",
    "ADSK_Размер_Толщина",
}

_SOURCE_CACHE = {}


def FormatValue(value, unit_type):
    if value is None:
        return None

    try:
        if unit_type == "length":
            mm_value = value * FEET_TO_MM
            return "{:.0f}".format(mm_value)

        elif unit_type == "area":
            m2_value = value * FEET2_TO_M2
            return "{:.3f}".format(m2_value)

        elif unit_type == "volume":
            m3_value = value * FEET3_TO_M3
            return "{:.4f}".format(m3_value)

        else:
            return str(value)
    except:
        return None


def GetBuiltInParamValue(element, bip):
    try:
        param = element.get_Parameter(bip)
        if param and param.HasValue:
            if param.StorageType == StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == StorageType.Integer:
                return float(param.AsInteger())
        return None
    except:
        return None


def GetCategoryBic(element):
    try:
        cat = element.Category
        if not cat:
            return None
        cat_id = cat.Id.IntegerValue
        for bic in MODEL_CATEGORIES:
            if int(bic) == cat_id:
                return bic
        return None
    except:
        return None


def GetDimensionValue(element, param_config):
    cat_bic = GetCategoryBic(element)
    if not cat_bic:
        return None

    # Для общестроительных категорий сразу считаем габариты по BoundingBox
    computed_value = GetGeneralConstructionDimensionValue(element, cat_bic, param_config)
    if computed_value is not None:
        return computed_value

    source_map = _BuildSourceMapForParam(param_config)
    bips = source_map.get(cat_bic, [])

    for bip in bips:
        value = GetBuiltInParamValue(element, bip)
        if value is not None:
            return value

    return None


def _BuildSourceMapForParam(param_config):
    param_name = param_config.get("NAME")
    if not param_name:
        return {}
    if param_name in _SOURCE_CACHE:
        return _SOURCE_CACHE[param_name]

    source_map = {}
    for source in param_config.get("SOURCES", []):
        bip = source.get("BIP")
        if not bip:
            continue
        for bic in source.get("CATEGORIES", []):
            if bic not in source_map:
                source_map[bic] = []
            source_map[bic].append(bip)

    _SOURCE_CACHE[param_name] = source_map
    return source_map


def _GetBoundingBoxSizes(element):
    try:
        bbox = element.get_BoundingBox(None)
        if not bbox:
            return None

        dx = abs(bbox.Max.X - bbox.Min.X)
        dy = abs(bbox.Max.Y - bbox.Min.Y)
        dz = abs(bbox.Max.Z - bbox.Min.Z)

        if dx <= 1e-9 and dy <= 1e-9 and dz <= 1e-9:
            return None

        return dx, dy, dz
    except Exception:
        return None


def GetGeneralConstructionDimensionValue(element, cat_bic, param_config):
    param_name = param_config.get("NAME")
    if param_name not in GENERAL_DIMENSION_PARAMETERS:
        return None
    if cat_bic not in GENERAL_CONSTRUCTION_CATEGORIES:
        return None

    sizes = _GetBoundingBoxSizes(element)
    if not sizes:
        return None

    dx, dy, dz = sizes
    horizontal_max = max(dx, dy)
    horizontal_min = min(dx, dy)

    if param_name == "ADSK_Размер_Длина":
        return horizontal_max
    if param_name == "ADSK_Размер_Ширина":
        return horizontal_min
    if param_name == "ADSK_Размер_Высота":
        return dz
    if param_name == "ADSK_Размер_Толщина":
        return min(dx, dy, dz)

    return None


def SetParameterStringValue(element, param_name, value):
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


def EnsureParametersExist(doc):
    from Autodesk.Revit.DB import BuiltInParameterGroup

    results = {
        "added": [],
        "existing": [],
        "failed": [],
    }

    for param_def in DIMENSION_PARAMETERS:
        param_config = {
            "PARAMETER_NAME": param_def["NAME"],
            "BINDING_TYPE": "Instance",
            "PARAMETER_GROUP": BuiltInParameterGroup.PG_GEOMETRY,
            "CATEGORIES": MODEL_CATEGORIES,
        }

        try:
            result = AddSharedParameterToDoc(doc, param_config)
            if result.get("success"):
                if result.get("mode") == "added":
                    results["added"].append(param_def["NAME"])
                else:
                    results["existing"].append(param_def["NAME"])
            else:
                results["failed"].append(param_def["NAME"])
        except Exception as e:
            results["failed"].append(param_def["NAME"])

    return results


def GetElementsToProcess(doc):
    collector = FilteredElementCollector(doc)
    cats = List[BuiltInCategory]()
    for bic in MODEL_CATEGORIES:
        cats.Add(bic)
    category_filter = ElementMulticategoryFilter(cats)
    return collector.WhereElementIsNotElementType().WherePasses(category_filter)


def FillDimensions(doc, progress_callback=None):
    _SOURCE_CACHE.clear()
    elements = GetElementsToProcess(doc)
    total = elements.GetElementCount()

    stats = {
        "total": total,
        "updated": 0,
        "skipped": 0,
    }

    param_stats = {}
    for param_def in DIMENSION_PARAMETERS:
        param_stats[param_def["NAME"]] = {
            "updated": 0,
            "skipped": 0,
            "no_value": 0,
        }

    if total == 0:
        return {
            "total": 0,
            "updated_count": 0,
            "skipped_count": 0,
            "param_stats": param_stats,
            "filled": False,
        }

    current_index = 0

    for element in elements:
        if progress_callback:
            progress = int(((current_index + 1) / float(total)) * 100)
            progress_callback(progress)

        element_updated = False

        for param_def in DIMENSION_PARAMETERS:
            param_name = param_def["NAME"]
            unit_type = param_def["UNIT_TYPE"]

            raw_value = GetDimensionValue(element, param_def)

            if raw_value is not None:
                formatted_value = FormatValue(raw_value, unit_type)

                if formatted_value:
                    result = SetParameterStringValue(
                        element, param_name, formatted_value
                    )

                    if result["status"] == "updated":
                        param_stats[param_name]["updated"] += 1
                        element_updated = True
                    elif result["status"] == "already_ok":
                        param_stats[param_name]["skipped"] += 1
                    else:
                        param_stats[param_name]["skipped"] += 1
                else:
                    param_stats[param_name]["no_value"] += 1
            else:
                param_stats[param_name]["no_value"] += 1

        if element_updated:
            stats["updated"] += 1
        else:
            stats["skipped"] += 1

        current_index += 1

    filled = stats["updated"] > 0

    return {
        "total": total,
        "updated_count": stats["updated"],
        "skipped_count": stats["skipped"],
        "param_stats": param_stats,
        "filled": filled,
    }


def Execute(doc, progress_callback=None):
    t = None

    try:
        if not doc.IsModifiable:
            t = Transaction(doc, "Заполнение параметров размеров ADSK")
            t.Start()

        param_results = EnsureParametersExist(doc)

        failed_count = len(param_results.get("failed", []))
        if failed_count == len(DIMENSION_PARAMETERS):
            if t is not None:
                t.RollBack()
            return {
                "success": False,
                "message": "Не удалось добавить ни одного параметра",
                "parameters": param_results,
                "fill": {
                    "filled": False,
                    "total": 0,
                    "updated_count": 0,
                    "skipped_count": 0,
                    "param_stats": {},
                },
            }

        fill_result = FillDimensions(doc, progress_callback)

        if t is not None:
            t.Commit()

        message = "Заполнение завершено"
        if fill_result["updated_count"] > 0:
            message = "Обновлено элементов: {0}".format(fill_result["updated_count"])
        else:
            message = "Нет элементов для обновления"

        return {
            "success": True,
            "parameters": param_results,
            "message": message,
            "fill": {
                "filled": fill_result["filled"],
                "total": fill_result["total"],
                "updated_count": fill_result["updated_count"],
                "skipped_count": fill_result["skipped_count"],
                "param_stats": fill_result["param_stats"],
            },
        }

    except Exception as e:
        if t is not None:
            t.RollBack()
        return {
            "success": False,
            "message": "Ошибка: {0}".format(str(e)),
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "filled": False,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "param_stats": {},
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
