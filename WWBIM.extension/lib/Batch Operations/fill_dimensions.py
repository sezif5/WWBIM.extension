# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import os
import inspect
import math

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
from System import Enum
from System.Collections.Generic import List

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


DIMENSION_PARAMETERS = [
    {
        "NAME": "ADSK_Размер_Объём",
        "UNIT_TYPE": "volume",
        "SOURCES": [
            {
                "BIP": "HOST_VOLUME_COMPUTED",
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_Ceilings,
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": "HOST_VOLUME_COMPUTED",
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
            {
                "BIP": "RBS_PIPE_VOLUME_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                ],
            },
            {
                "BIP": "RBS_DUCT_VOLUME_PARAM",
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
                "BIP": "CURVE_ELEM_LENGTH",
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Ramps,
                    BuiltInCategory.OST_CurtainWallMullions,
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
            {
                "BIP": "STRUCTURAL_FRAME_CUT_LENGTH",
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralFraming,
                ],
            },
            {
                "BIP": "INSTANCE_LENGTH_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": "RBS_PIPE_LENGTH_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                ],
            },
            {
                "BIP": "RBS_DUCT_LENGTH_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                ],
            },
            {
                "BIP": "RBS_CABLETRAY_LENGTH_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_CableTray,
                ],
            },
            {
                "BIP": "RBS_CONDUIT_LENGTH_PARAM",
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
                "BIP": "WALL_USER_WIDTH_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                ],
            },
            {
                "BIP": "FLOOR_ATTR_THICKNESS_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_Floors,
                ],
            },
            {
                "BIP": "CEILING_THICKNESS",
                "CATEGORIES": [
                    BuiltInCategory.OST_Ceilings,
                ],
            },
            {
                "BIP": "STRUCTURAL_FOUNDATION_THICKNESS",
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": "FAMILY_WIDTH_PARAM",
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
                "BIP": "RBS_DUCT_WIDTH_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctAccessory,
                ],
            },
            {
                "BIP": "RBS_CABLETRAY_WIDTH_PARAM",
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
                "BIP": "FAMILY_HEIGHT_PARAM",
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
                "BIP": "RBS_DUCT_HEIGHT_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_FlexDuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctAccessory,
                ],
            },
            {
                "BIP": "RBS_CABLETRAY_HEIGHT_PARAM",
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
                "BIP": "WALL_USER_WIDTH_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                ],
            },
            {
                "BIP": "FLOOR_ATTR_THICKNESS_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_Floors,
                ],
            },
            {
                "BIP": "CEILING_THICKNESS",
                "CATEGORIES": [
                    BuiltInCategory.OST_Ceilings,
                ],
            },
            {
                "BIP": "STRUCTURAL_FOUNDATION_THICKNESS",
                "CATEGORIES": [
                    BuiltInCategory.OST_StructuralFoundation,
                ],
            },
            {
                "BIP": "FAMILY_DEPTH_PARAM",
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
                "BIP": "HOST_AREA_COMPUTED",
                "CATEGORIES": [
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_Ceilings,
                ],
            },
            {
                "BIP": "STRUCTURAL_AREA",
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
                "BIP": "RBS_PIPE_DIAMETER_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory,
                ],
            },
            {
                "BIP": "RBS_PIPE_OUTER_DIAMETER_PARAM",
                "CATEGORIES": [
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_FlexPipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory,
                ],
            },
            {
                "BIP": "RBS_DUCT_DIAMETER_PARAM",
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

try:
    STRING_TYPES = (basestring,)
except NameError:
    STRING_TYPES = (str,)


def _resolve_bip(name):
    if not name:
        return None
    try:
        if Enum.IsDefined(BuiltInParameter, name):
            return Enum.Parse(BuiltInParameter, name)
    except Exception:
        pass
    return None


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


def _GetElementHeightFromBoundingBox(element):
    try:
        bbox = element.get_BoundingBox(None)
        if not bbox:
            return None
        dz = abs(bbox.Max.Z - bbox.Min.Z)
        if dz <= 1e-9:
            return None
        return dz
    except Exception:
        return None


def _GetWallDimensionValue(element, param_name):
    if param_name == "ADSK_Размер_Ширина":
        return None

    if param_name == "ADSK_Размер_Высота":
        return _GetElementHeightFromBoundingBox(element)

    if param_name == "ADSK_Размер_Длина":
        try:
            loc = element.Location
            if loc and hasattr(loc, "Curve") and loc.Curve:
                length = loc.Curve.Length
                if length is not None and length > 1e-9:
                    return length
        except Exception:
            pass

        length_bip = _resolve_bip("CURVE_ELEM_LENGTH")
        if length_bip is not None:
            return GetBuiltInParamValue(element, length_bip)
        return None

    if param_name == "ADSK_Размер_Толщина":
        try:
            wall_type = element.WallType
            if wall_type and wall_type.Width is not None and wall_type.Width > 1e-9:
                return wall_type.Width
        except Exception:
            pass

        width_bip = _resolve_bip("WALL_USER_WIDTH_PARAM")
        if width_bip is not None:
            return GetBuiltInParamValue(element, width_bip)
        return None

    return None


def _GetFloorDimensionValue(element, param_name):
    """Получение размеров для перекрытий из системных параметров."""
    if param_name == "ADSK_Размер_Толщина":
        # Сначала пытаемся получить толщину из типа перекрытия
        try:
            floor_type = element.FloorType
            if floor_type:
                # Пробуем получить толщину через CompoundStructure
                try:
                    cs = floor_type.GetCompoundStructure()
                    if cs and cs.Width is not None and cs.Width > 1e-9:
                        return cs.Width
                except Exception:
                    pass
        except Exception:
            pass

        # Fallback на параметр типа
        thickness_bip = _resolve_bip("FLOOR_ATTR_THICKNESS_PARAM")
        if thickness_bip is not None:
            try:
                floor_type = element.FloorType
                if floor_type:
                    param = floor_type.get_Parameter(thickness_bip)
                    if param and param.HasValue:
                        if param.StorageType == StorageType.Double:
                            return param.AsDouble()
            except Exception:
                pass
        return None

    return None


# Линейные инженерные категории
MEP_LINEAR_CATEGORIES = {
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_Conduit,
}


def _GetMepLinearDimensionValue(element, cat_bic, param_name):
    """Получение размеров для линейных инженерных категорий из системных параметров."""

    # Трубы (PipeCurves, FlexPipeCurves)
    if cat_bic in (BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_FlexPipeCurves):
        if param_name == "ADSK_Размер_Длина":
            bip = _resolve_bip("RBS_PIPE_LENGTH_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        elif param_name == "ADSK_Размер_Диаметр":
            # Сначала внешний диаметр, потом обычный
            bip = _resolve_bip("RBS_PIPE_OUTER_DIAMETER_PARAM")
            if bip:
                value = GetBuiltInParamValue(element, bip)
                if value is not None:
                    return value
            bip = _resolve_bip("RBS_PIPE_DIAMETER_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        return None

    # Воздуховоды (DuctCurves, FlexDuctCurves)
    if cat_bic in (BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_FlexDuctCurves):
        if param_name == "ADSK_Размер_Длина":
            bip = _resolve_bip("RBS_DUCT_LENGTH_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        elif param_name == "ADSK_Размер_Ширина":
            bip = _resolve_bip("RBS_DUCT_WIDTH_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        elif param_name == "ADSK_Размер_Высота":
            bip = _resolve_bip("RBS_DUCT_HEIGHT_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        elif param_name == "ADSK_Размер_Диаметр":
            bip = _resolve_bip("RBS_DUCT_DIAMETER_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        return None

    # Кабельные лотки (CableTray)
    if cat_bic == BuiltInCategory.OST_CableTray:
        if param_name == "ADSK_Размер_Длина":
            bip = _resolve_bip("RBS_CABLETRAY_LENGTH_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        elif param_name == "ADSK_Размер_Ширина":
            bip = _resolve_bip("RBS_CABLETRAY_WIDTH_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        elif param_name == "ADSK_Размер_Высота":
            bip = _resolve_bip("RBS_CABLETRAY_HEIGHT_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        return None

    # Кабель-каналы (Conduit)
    if cat_bic == BuiltInCategory.OST_Conduit:
        if param_name == "ADSK_Размер_Длина":
            bip = _resolve_bip("RBS_CONDUIT_LENGTH_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        elif param_name == "ADSK_Размер_Диаметр":
            bip = _resolve_bip("RBS_CONDUIT_DIAMETER_PARAM")
            if bip:
                return GetBuiltInParamValue(element, bip)
        return None

    return None


def GetDimensionValue(element, param_config):
    cat_bic = GetCategoryBic(element)
    if not cat_bic:
        return None

    param_name = param_config.get("NAME")

    # Параметры Объём, Площадь, Толщина - только для общестроительных категорий
    CONSTRUCTION_ONLY_PARAMS = {
        "ADSK_Размер_Объём",
        "ADSK_Площадь",
        "ADSK_Размер_Толщина",
    }
    if (
        param_name in CONSTRUCTION_ONLY_PARAMS
        and cat_bic not in GENERAL_CONSTRUCTION_CATEGORIES
    ):
        return None

    if (
        cat_bic == BuiltInCategory.OST_Walls
        and param_name in GENERAL_DIMENSION_PARAMETERS
    ):
        return _GetWallDimensionValue(element, param_name)

    # Для перекрытий - сначала системный параметр толщины
    if cat_bic == BuiltInCategory.OST_Floors and param_name == "ADSK_Размер_Толщина":
        floor_value = _GetFloorDimensionValue(element, param_name)
        if floor_value is not None:
            return floor_value

    # Для линейных инженерных категорий - системные параметры
    if cat_bic in MEP_LINEAR_CATEGORIES:
        mep_value = _GetMepLinearDimensionValue(element, cat_bic, param_name)
        if mep_value is not None:
            return mep_value

    source_map = _BuildSourceMapForParam(param_config)
    bips = source_map.get(cat_bic, [])

    for bip in bips:
        value = GetBuiltInParamValue(element, bip)
        if value is not None:
            return value

    # Для общестроительных категорий используем геометрию только как fallback,
    # если системные параметры не дали значения.
    computed_value = GetGeneralConstructionDimensionValue(
        element, cat_bic, param_config
    )
    if computed_value is not None:
        return computed_value

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

        if isinstance(bip, STRING_TYPES):
            resolved_bip = _resolve_bip(bip)
            if resolved_bip is None:
                continue
            source["BIP"] = resolved_bip
            bip = resolved_bip

        for bic in source.get("CATEGORIES", []):
            if bic not in source_map:
                source_map[bic] = []
            source_map[bic].append(bip)

    _SOURCE_CACHE[param_name] = source_map
    return source_map


def _GetSolidLocalBoundingBoxSizes(element):
    try:
        from Autodesk.Revit.DB import Options, Solid, GeometryInstance, Transform, XYZ

        def _safe_normalize(vec):
            if not vec:
                return None
            try:
                length = vec.GetLength()
            except Exception:
                return None
            if length <= 1e-9:
                return None
            return vec.Normalize()

        def _get_element_axes(el):
            z_axis = XYZ.BasisZ
            x_axis = None

            try:
                loc = el.Location
                if loc and hasattr(loc, "Curve") and loc.Curve:
                    p0 = loc.Curve.GetEndPoint(0)
                    p1 = loc.Curve.GetEndPoint(1)
                    curve_dir = _safe_normalize(p1 - p0)
                    if curve_dir and abs(curve_dir.DotProduct(z_axis)) < 0.95:
                        x_axis = curve_dir
            except Exception:
                pass

            if x_axis is None:
                try:
                    loc = el.Location
                    if loc and hasattr(loc, "Rotation"):
                        angle = loc.Rotation
                        x_axis = _safe_normalize(
                            XYZ(math.cos(angle), math.sin(angle), 0.0)
                        )
                except Exception:
                    pass

            if x_axis is None:
                try:
                    t = el.GetTransform()
                    x_axis = _safe_normalize(t.BasisX)
                except Exception:
                    x_axis = None

            if x_axis is None:
                try:
                    hand = _safe_normalize(el.HandOrientation)
                    if hand:
                        x_axis = hand
                except Exception:
                    pass

            if x_axis is None:
                try:
                    facing = _safe_normalize(el.FacingOrientation)
                    if facing:
                        x_axis = facing
                except Exception:
                    pass

            if x_axis is None:
                x_axis = XYZ.BasisX

            y_axis = z_axis.CrossProduct(x_axis)
            y_axis = _safe_normalize(y_axis)

            if y_axis is None:
                x_axis = XYZ.BasisX
                y_axis = XYZ.BasisY
                z_axis = XYZ.BasisZ
                return x_axis, y_axis, z_axis

            x_axis = _safe_normalize(y_axis.CrossProduct(z_axis))
            if x_axis is None:
                x_axis = XYZ.BasisX
                y_axis = XYZ.BasisY
                z_axis = XYZ.BasisZ
                return x_axis, y_axis, z_axis

            return x_axis, y_axis, z_axis

        def _iter_solids(geo_element):
            if not geo_element:
                return
            for geo_obj in geo_element:
                if isinstance(geo_obj, Solid):
                    try:
                        if geo_obj.Volume > 1e-9 and geo_obj.Edges.Size > 0:
                            yield geo_obj
                    except Exception:
                        pass
                    continue

                if isinstance(geo_obj, GeometryInstance):
                    try:
                        inst_geo = geo_obj.GetInstanceGeometry()
                    except Exception:
                        inst_geo = None
                    if not inst_geo:
                        continue

                    for nested in _iter_solids(inst_geo):
                        yield nested

        geo = element.get_Geometry(Options())
        if not geo:
            return None

        x_axis, y_axis, z_axis = _get_element_axes(element)

        min_x = None
        max_x = None
        min_y = None
        max_y = None
        min_z = None
        max_z = None

        for solid in _iter_solids(geo):
            try:
                solid_bbox = solid.GetBoundingBox()
            except Exception:
                solid_bbox = None
            if not solid_bbox:
                continue

            try:
                bb_tr = solid_bbox.Transform
            except Exception:
                bb_tr = Transform.Identity

            pmin = solid_bbox.Min
            pmax = solid_bbox.Max

            corners = (
                XYZ(pmin.X, pmin.Y, pmin.Z),
                XYZ(pmin.X, pmin.Y, pmax.Z),
                XYZ(pmin.X, pmax.Y, pmin.Z),
                XYZ(pmin.X, pmax.Y, pmax.Z),
                XYZ(pmax.X, pmin.Y, pmin.Z),
                XYZ(pmax.X, pmin.Y, pmax.Z),
                XYZ(pmax.X, pmax.Y, pmin.Z),
                XYZ(pmax.X, pmax.Y, pmax.Z),
            )

            for point in corners:
                try:
                    world_point = bb_tr.OfPoint(point)
                except Exception:
                    world_point = point

                vx = world_point.DotProduct(x_axis)
                vy = world_point.DotProduct(y_axis)
                vz = world_point.DotProduct(z_axis)

                if min_x is None or vx < min_x:
                    min_x = vx
                if max_x is None or vx > max_x:
                    max_x = vx
                if min_y is None or vy < min_y:
                    min_y = vy
                if max_y is None or vy > max_y:
                    max_y = vy
                if min_z is None or vz < min_z:
                    min_z = vz
                if max_z is None or vz > max_z:
                    max_z = vz

        if (
            min_x is None
            or max_x is None
            or min_y is None
            or max_y is None
            or min_z is None
            or max_z is None
        ):
            return None

        dx = abs(max_x - min_x)
        dy = abs(max_y - min_y)
        dz = abs(max_z - min_z)

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

    sizes = _GetSolidLocalBoundingBoxSizes(element)
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


def SetParameterValue(element, param_name, raw_value, formatted_value):
    """Установка значения параметра с поддержкой String и Double типов."""
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return {"status": "parameter_not_found", "reason": "parameter_not_found"}

        if param.IsReadOnly:
            return {"status": "readonly", "reason": "readonly"}

        storage_type = param.StorageType

        # Для строковых параметров - форматируемое значение
        if storage_type == StorageType.String:
            try:
                current_value = param.AsString()
                if current_value == formatted_value:
                    return {"status": "already_ok", "reason": "already_ok"}
            except:
                pass

            if formatted_value is not None:
                param.Set(formatted_value)
                return {"status": "updated", "reason": None}

            return {"status": "exception", "reason": "value_is_none"}

        # Для числовых параметров (Double) - raw значение в feet
        if storage_type == StorageType.Double:
            if raw_value is None:
                return {"status": "exception", "reason": "value_is_none"}

            try:
                current_value = param.AsDouble()
                # Сравниваем с допуском
                if abs(current_value - raw_value) < 1e-9:
                    return {"status": "already_ok", "reason": "already_ok"}
            except:
                pass

            param.Set(raw_value)
            return {"status": "updated", "reason": None}

        return {"status": "wrong_storage_type", "reason": "unsupported_storage_type"}
    except Exception as e:
        return {"status": "exception", "reason": "exception"}


def EnsureParametersExist(doc):
    from Autodesk.Revit.DB import BuiltInParameterGroup

    # Параметры только для общестроительных категорий
    CONSTRUCTION_ONLY_PARAMS = {
        "ADSK_Размер_Объём",
        "ADSK_Площадь",
        "ADSK_Размер_Толщина",
    }

    results = {
        "added": [],
        "existing": [],
        "failed": [],
        "failed_details": [],
    }

    for param_def in DIMENSION_PARAMETERS:
        param_name = param_def["NAME"]

        # Выбираем категории в зависимости от параметра
        if param_name in CONSTRUCTION_ONLY_PARAMS:
            target_categories = GENERAL_CONSTRUCTION_CATEGORIES
        else:
            target_categories = MODEL_CATEGORIES

        param_config = {
            "PARAMETER_NAME": param_name,
            "BINDING_TYPE": "Instance",
            "PARAMETER_GROUP": BuiltInParameterGroup.PG_GEOMETRY,
            "CATEGORIES": target_categories,
        }

        try:
            result = AddSharedParameterToDoc(doc, param_config)
            if result.get("success"):
                if result.get("mode") == "added":
                    results["added"].append(param_def["NAME"])
                else:
                    results["existing"].append(param_def["NAME"])
            else:
                failed_name = param_def["NAME"]
                results["failed"].append(failed_name)
                results["failed_details"].append(
                    {
                        "name": failed_name,
                        "message": result.get("message", "Неизвестная ошибка"),
                        "mode": result.get("mode"),
                        "diagnostics": result.get("diagnostics", {}),
                    }
                )
        except Exception as e:
            failed_name = param_def["NAME"]
            results["failed"].append(failed_name)
            results["failed_details"].append(
                {
                    "name": failed_name,
                    "message": "Exception: {0}".format(str(e)),
                    "mode": "exception",
                    "diagnostics": {},
                }
            )

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

                result = SetParameterValue(
                    element, param_name, raw_value, formatted_value
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

            failed_details = param_results.get("failed_details", [])
            if failed_details:
                details_text = "; ".join(
                    [
                        "{0}: {1}".format(
                            item.get("name", "?"), item.get("message", "")
                        )
                        for item in failed_details
                    ]
                )
                fail_message = "Не удалось добавить ни одного параметра. {0}".format(
                    details_text
                )
            else:
                fail_message = "Не удалось добавить ни одного параметра"

            return {
                "success": False,
                "message": fail_message,
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

        failed_details = param_results.get("failed_details", [])
        if failed_details:
            details_text = "; ".join(
                [
                    "{0}: {1}".format(item.get("name", "?"), item.get("message", ""))
                    for item in failed_details
                ]
            )
            message = "{0}. Ошибки добавления параметров ({1}): {2}".format(
                message, len(failed_details), details_text
            )

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
