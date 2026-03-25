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
    ScheduleSheetInstance,
    ViewSheet,
    View,
    Parameter,
    StorageType,
    ElementId,
    ViewType,
    FamilyInstance,
    Transaction,
    LabelUtils,
)
import System
from System.Collections.Generic import List

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


CONFIG = {
    "PARAMETER_NAME": "ADSK_КомплектШифр",
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


def DetermineKitMode(doc):
    filename = doc.Title.upper()
    prefixes = ["_AR", "_AI", "_НАВ", "_KM", "_KR", "_KG"]

    for prefix in prefixes:
        if prefix in filename:
            return "Schedules"

    return "Views"


def SafeParamElemName(param_elem):
    try:
        if hasattr(param_elem, "GetDefinition"):
            definition = param_elem.GetDefinition()
            if definition:
                return definition.Name
    except:
        pass
    try:
        return param_elem.Name
    except:
        return None


def GetBrowserOrganizationSheetParameters(doc):
    debug_info = {"params": [], "errors": []}
    try:
        from Autodesk.Revit.DB import BrowserOrganization, ViewSheet

        org = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc)
        if not org:
            debug_info["errors"].append(
                "GetCurrentBrowserOrganizationForSheets returned None"
            )
            return [], debug_info

        param_id_to_name = {}

        sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
        debug_info["params"].append("Total sheets: {0}".format(len(sheets)))

        for sheet in sheets:
            try:
                items = org.GetFolderItems(sheet.Id)
                if items and len(items) > 0:
                    for info in items:
                        try:
                            if info:
                                group_header = info.Name
                                element_id = info.ElementId

                                if element_id and element_id.IntegerValue > 0:
                                    if element_id not in param_id_to_name:
                                        param_id_to_name[element_id] = group_header

                                        param_elem = doc.GetElement(element_id)
                                        if param_elem:
                                            name = SafeParamElemName(param_elem)
                                            if name:
                                                param_id_to_name[element_id] = name
                                            else:
                                                debug_info["errors"].append(
                                                    "ElementId {0}: SafeParamElemName returned None".format(
                                                        element_id.IntegerValue
                                                    )
                                                )
                                        else:
                                            try:
                                                builtin_param = BuiltInParameter(
                                                    element_id.IntegerValue
                                                )
                                                name = LabelUtils.GetLabelFor(
                                                    builtin_param
                                                )
                                                if name:
                                                    param_id_to_name[element_id] = name
                                            except:
                                                debug_info["errors"].append(
                                                    "ElementId {0}: Not a BuiltInParameter".format(
                                                        element_id.IntegerValue
                                                    )
                                                )
                        except Exception as inner_e:
                            debug_info["errors"].append(
                                "sheet {0}, item: {1}".format(
                                    sheet.Id.IntegerValue, str(inner_e)
                                )
                            )
            except Exception as e:
                debug_info["errors"].append(
                    "sheet {0}: GetFolderItems failed: {1}".format(
                        sheet.Id.IntegerValue, str(e)
                    )
                )

        param_names = list(param_id_to_name.values())
        debug_info["params"].append(
            "Unique param_ids found: {0}".format(len(param_id_to_name))
        )
        debug_info["params"].append("Param names: {0}".format(", ".join(param_names)))
        return param_names, debug_info
    except Exception as e:
        debug_info["errors"].append("outer: {0}".format(str(e)))
        return [], debug_info


def IsValidKitCodeValue(value):
    kit_codes = [
        "АР",
        "КЖ",
        "ОВ",
        "ВК",
        "ИТП",
        "ВНС",
        "АИ",
        "СС",
        "ЭОМ",
        "КМ",
        "КК",
        "СВН",
        "ЛВС",
        "КР",
    ]
    if not value:
        return False
    value_upper = value.upper()
    for code in kit_codes:
        if code in value.upper():
            return True
    return False


def GetSheetsWithValidKitCodes(doc, sheet_param_name, kit_mode):
    if not sheet_param_name:
        return set()

    sheets_to_check = []
    if kit_mode == "Schedules":
        schedule_sheet_instances = (
            FilteredElementCollector(doc).OfClass(ScheduleSheetInstance).ToElements()
        )
        sheets_with_schedules = set()
        for inst in schedule_sheet_instances:
            sheet_id = None
            if hasattr(inst, "OwnerSheetId"):
                sheet_id = inst.OwnerSheetId
            elif hasattr(inst, "OwnerViewId"):
                sheet_id = inst.OwnerViewId
            if sheet_id and sheet_id.IntegerValue > 0:
                sheets_with_schedules.add(sheet_id.IntegerValue)

        for sheet_id_int in sheets_with_schedules:
            sheet_id = ElementId(sheet_id_int)
            sheet = doc.GetElement(sheet_id)
            if sheet:
                sheets_to_check.append(sheet)
    else:
        sheets_to_check = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

    valid_sheets = set()
    for sheet in sheets_to_check:
        param = sheet.LookupParameter(sheet_param_name)
        if param and param.HasValue:
            try:
                value = param.AsString()
                if IsValidKitCodeValue(value):
                    valid_sheets.add(sheet.Id.IntegerValue)
            except:
                pass

    return valid_sheets


def DetermineSheetParameterName(doc, kit_mode):
    param_names, browser_debug = GetBrowserOrganizationSheetParameters(doc)
    debug_error = None

    sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

    if param_names:
        for param_name in param_names:
            valid_sheets = GetSheetsWithValidKitCodes(doc, param_name, kit_mode)
            if len(valid_sheets) > 0:
                return {
                    "name": param_name,
                    "source": "browser_organization_group_by",
                    "all_browser_params": param_names,
                    "debug_info": browser_debug,
                    "debug_error": None,
                }

    priority_params = [
        "ADSK_Штамп Раздел проекта",
        "ADSK_Раздел проекта",
        "Раздел проекта",
    ]

    for param_name in priority_params:
        values = set()
        for sheet in sheets:
            param = sheet.LookupParameter(param_name)
            if param and param.HasValue:
                try:
                    values.add(param.AsString())
                except:
                    pass
        if len(values) > 0:
            return {
                "name": param_name,
                "source": "fallback_priority_list",
                "all_browser_params": param_names,
                "debug_info": browser_debug,
                "debug_error": debug_error,
            }

    error_parts = []
    if param_names:
        error_parts.append("Browser params: {}".format(", ".join(param_names)))
        error_parts.append("Params checked: {}".format(len(param_names)))
    else:
        error_parts.append("No browser params found")

    sheets_checked = len(sheets)
    error_parts.append("Sheets checked: {}".format(sheets_checked))

    if browser_debug:
        if browser_debug.get("params"):
            error_parts.append("Debug: {}".format("; ".join(browser_debug["params"])))
        if browser_debug.get("errors"):
            error_parts.append("Errors: {}".format("; ".join(browser_debug["errors"])))

    debug_error = "; ".join(error_parts)

    return {
        "name": None,
        "source": "not_found",
        "all_browser_params": param_names,
        "debug_info": browser_debug,
        "debug_error": debug_error,
    }


def GetSheetParameter(sheet, param_name):
    param = sheet.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsString()
    return None


def GetCategoryFromSchedule(scheduleInstance):
    try:
        schedule_id = None
        if hasattr(scheduleInstance, "ScheduleId"):
            schedule_id = scheduleInstance.ScheduleId
        elif hasattr(scheduleInstance, "ViewScheduleId"):
            schedule_id = scheduleInstance.ViewScheduleId

        if schedule_id and schedule_id.IntegerValue > 0:
            return schedule_id
        return None
    except:
        return None


def GetElementsFromSchedule(doc, scheduleInstance):
    debug = {
        "schedule_id": None,
        "cat_id": None,
        "used_collector_view": False,
        "used_category_ids": False,
        "category_count": 0,
        "elements_count": 0,
        "errors": [],
    }
    try:
        elements = []

        schedule_id = None
        if hasattr(scheduleInstance, "ScheduleId"):
            schedule_id = scheduleInstance.ScheduleId
        elif hasattr(scheduleInstance, "ViewScheduleId"):
            schedule_id = scheduleInstance.ViewScheduleId

        debug["schedule_id"] = schedule_id.IntegerValue if schedule_id else None

        if not schedule_id or schedule_id.IntegerValue <= 0:
            debug["errors"].append("invalid schedule_id")
            return [], debug

        schedule_view = doc.GetElement(schedule_id)
        if not schedule_view:
            debug["errors"].append("schedule_view not found")
            return [], debug

        cat_id = None
        if hasattr(schedule_view, "Definition"):
            try:
                cat_id = schedule_view.Definition.CategoryId
            except:
                pass

        debug["cat_id"] = cat_id.IntegerValue if cat_id else None

        if cat_id and cat_id.IntegerValue > 0:
            collector = FilteredElementCollector(doc)
            elements = (
                collector.OfCategoryId(cat_id)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        else:
            try:
                debug["used_collector_view"] = True
                collector = FilteredElementCollector(doc, schedule_view.Id)
                elements = collector.WhereElementIsNotElementType().ToElements()
                debug["elements_count"] = len(elements)
                if elements:
                    return elements, debug
            except Exception as e:
                debug["errors"].append("collector_view: {0}".format(str(e)))

            try:
                if hasattr(schedule_view, "Definition"):
                    def_obj = schedule_view.Definition
                    if def_obj and hasattr(def_obj, "GetCategoryIds"):
                        cat_ids = def_obj.GetCategoryIds()
                        if cat_ids and cat_ids.Count > 0:
                            debug["used_category_ids"] = True
                            debug["category_count"] = cat_ids.Count
                            categories_list = list(cat_ids)
                            if cat_ids.Count == 1:
                                collector = FilteredElementCollector(doc)
                                elements = (
                                    collector.OfCategoryId(categories_list[0])
                                    .WhereElementIsNotElementType()
                                    .ToElements()
                                )
                            else:
                                elements = []
                                for cat in categories_list:
                                    collector = FilteredElementCollector(doc)
                                    cat_elems = (
                                        collector.OfCategoryId(cat)
                                        .WhereElementIsNotElementType()
                                        .ToElements()
                                    )
                                    elements.extend(cat_elems)
                            if elements:
                                debug["elements_count"] = len(elements)
                                return elements, debug
            except Exception as e:
                debug["errors"].append("category_ids: {0}".format(str(e)))

        debug["elements_count"] = len(elements)
        return elements, debug
    except Exception as e:
        debug["errors"].append("outer: {0}".format(str(e)))
        return [], debug


def GetElementsFromView(doc, view):
    try:
        collector = FilteredElementCollector(doc, view.Id)
        return collector.WhereElementIsNotElementType().ToElements()
    except:
        return []


def GetKitCodeParameter(element):
    param = element.LookupParameter("ADSK_КомплектШифр")
    if param and param.StorageType == StorageType.String:
        try:
            return param.AsString()
        except:
            return None
    return None


def SetKitCodeParameter(element, kit_code):
    try:
        param = element.LookupParameter("ADSK_КомплектШифр")
        if not param:
            return {"status": "parameter_not_found", "reason": "parameter_not_found"}

        if param.StorageType != StorageType.String:
            return {"status": "wrong_storage_type", "reason": "wrong_storage_type"}

        if param.IsReadOnly:
            return {"status": "readonly", "reason": "readonly"}

        # Если у элемента уже есть значение — не перезаписываем (первый выигрывает)
        if param.HasValue:
            existing_value = param.AsString()
            if existing_value:
                return {"status": "already_ok", "reason": "already_ok"}

        try:
            param.Set(kit_code)
            return {"status": "updated", "reason": None}
        except:
            pass

        return {"status": "exception", "reason": "value_is_none"}
    except Exception as e:
        return {"status": "exception", "reason": "exception"}


def FillKitCodes_Schedules(doc, sheet_param_name, progress_callback=None):
    schedule_instances = (
        FilteredElementCollector(doc).OfClass(ScheduleSheetInstance).ToElements()
    )

    total = 0
    updated_count = 0
    skipped_count = 0
    skip_reasons = {
        "parameter_not_found": 0,
        "readonly": 0,
        "wrong_storage_type": 0,
        "already_ok": 0,
        "exception": 0,
        "invalid_kit_code": 0,
    }
    all_values = set()

    total_schedules = len(schedule_instances)
    current_schedule = 0

    for schedule_inst in schedule_instances:
        if hasattr(schedule_inst, "OwnerSheetId"):
            sheet_id = schedule_inst.OwnerSheetId
        else:
            sheet_id = schedule_inst.OwnerViewId
        if sheet_id.IntegerValue == -1:
            continue

        sheet = doc.GetElement(sheet_id)
        if not sheet:
            continue

        kit_code = GetSheetParameter(sheet, sheet_param_name)
        if not kit_code:
            continue

        if not IsValidKitCodeValue(kit_code):
            continue

        all_values.add(kit_code)
        elements, debug_info = GetElementsFromSchedule(doc, schedule_inst)
        total += len(elements)

    current_schedule = 0
    first_debug = None

    for schedule_inst in schedule_instances:
        if progress_callback:
            progress = int((current_schedule / total_schedules) * 100)
            progress_callback(progress)

        if hasattr(schedule_inst, "OwnerSheetId"):
            sheet_id = schedule_inst.OwnerSheetId
        else:
            sheet_id = schedule_inst.OwnerViewId
        if sheet_id.IntegerValue == -1:
            current_schedule += 1
            continue

        sheet = doc.GetElement(sheet_id)
        if not sheet:
            current_schedule += 1
            continue

        kit_code = GetSheetParameter(sheet, sheet_param_name)
        if not kit_code:
            current_schedule += 1
            continue

        if not IsValidKitCodeValue(kit_code):
            current_schedule += 1
            continue

        elements, debug_info = GetElementsFromSchedule(doc, schedule_inst)
        if first_debug is None:
            first_debug = debug_info

        for element in elements:
            result = SetKitCodeParameter(element, kit_code)
            if result["status"] == "updated":
                updated_count += 1
            else:
                skipped_count += 1
                reason = result["reason"]
                if reason in skip_reasons:
                    skip_reasons[reason] += 1

        current_schedule += 1

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
        "total": total,
        "updated_count": updated_count,
        "skipped_count": skipped_count,
        "skip_reasons": skip_reasons,
        "values": sorted(list(all_values)),
        "filled": filled,
        "message": message,
        "debug_schedule": first_debug,
    }


def FillKitCodes_Views(doc, sheet_param_name, progress_callback=None):
    view_types = [
        ViewType.FloorPlan,
        ViewType.CeilingPlan,
        ViewType.Section,
        ViewType.Elevation,
        ViewType.ThreeD,
    ]

    sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

    total = 0
    updated_count = 0
    skipped_count = 0
    skip_reasons = {
        "parameter_not_found": 0,
        "readonly": 0,
        "wrong_storage_type": 0,
        "already_ok": 0,
        "exception": 0,
        "invalid_kit_code": 0,
    }
    all_values = set()

    for sheet in sheets:
        kit_code = GetSheetParameter(sheet, sheet_param_name)
        if not kit_code:
            continue

        if not IsValidKitCodeValue(kit_code):
            continue

        all_values.add(kit_code)
        view_ids = sheet.GetAllPlacedViews()

        for view_id in view_ids:
            view = doc.GetElement(view_id)
            if not view:
                continue

            if view.ViewType in view_types:
                elements = GetElementsFromView(doc, view)
                total += len(elements)

    for sheet in sheets:
        kit_code = GetSheetParameter(sheet, sheet_param_name)
        if not kit_code:
            continue

        if not IsValidKitCodeValue(kit_code):
            continue

        view_ids = sheet.GetAllPlacedViews()

        for view_id in view_ids:
            view = doc.GetElement(view_id)
            if not view:
                continue

            if view.ViewType in view_types:
                elements = GetElementsFromView(doc, view)

                for element in elements:
                    result = SetKitCodeParameter(element, kit_code)
                    if result["status"] == "updated":
                        updated_count += 1
                    else:
                        skipped_count += 1
                        reason = result["reason"]
                        if reason in skip_reasons:
                            skip_reasons[reason] += 1

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
        "total": total,
        "updated_count": updated_count,
        "skipped_count": skipped_count,
        "skip_reasons": skip_reasons,
        "values": sorted(list(all_values)),
        "filled": filled,
        "message": message,
        "debug_schedule": None,
    }


def Execute(doc, progress_callback=None):
    kit_mode = DetermineKitMode(doc)
    sheet_param_info = DetermineSheetParameterName(doc, kit_mode)
    sheet_param_name = sheet_param_info["name"]
    param_source = sheet_param_info["source"]
    debug_error = sheet_param_info.get("debug_error")
    all_browser_params = sheet_param_info.get("all_browser_params", [])
    debug_info = sheet_param_info.get("debug_info", None)

    if not sheet_param_name:
        return {
            "success": False,
            "message": "Не удалось определить параметр листа для шифра комплекта",
            "parameters": {"added": [], "existing": [], "failed": []},
            "info": {
                "doc_title": doc.Title,
                "kit_mode": kit_mode,
                "sheet_param_name": sheet_param_name,
                "param_source": param_source,
                "all_browser_params": all_browser_params,
                "debug_info": debug_info,
                "debug_error": debug_error,
            },
            "fill": {
                "target_param": "ADSK_КомплектШифр",
                "source_param": sheet_param_name,
                "filled": False,
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
                "message": "Не удалось определить параметр листа",
            },
        }

    t = Transaction(doc, "Заполнение параметра ADSK_КомплектШифр")
    t.Start()

    param_result = {
        "success": False,
        "parameters": {"added": [], "existing": [], "failed": []},
        "message": "Параметр не был добавлен",
    }

    try:
        param_result = EnsureParameterExists(doc)

        # Проверяем результат добавления параметра
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
                    "info": {
                        "doc_title": doc.Title,
                        "kit_mode": kit_mode,
                        "sheet_param_name": sheet_param_name,
                        "param_source": param_source,
                        "all_browser_params": all_browser_params,
                        "debug_info": debug_info,
                        "debug_error": debug_error,
                    },
                    "fill": {
                        "target_param": "ADSK_КомплектШифр",
                        "source_param": sheet_param_name,
                        "filled": False,
                        "planned_value": None,
                        "total": 0,
                        "updated_count": 0,
                        "skipped_count": 0,
                        "skip_reasons": {},
                        "values": [],
                        "message": "Не удалось добавить параметр",
                    },
                }

        if kit_mode == "Schedules":
            fill_result = FillKitCodes_Schedules(
                doc, sheet_param_name, progress_callback
            )
        else:
            fill_result = FillKitCodes_Views(doc, sheet_param_name, progress_callback)

        t.Commit()

        if fill_result.get("debug_schedule") and fill_result["total"] == 0:
            ds = fill_result["debug_schedule"]
            debug_error = "schedule_debug: cat_id={0}, schedule_id={1}, used_table_data={2}, table_rows={3}, table_element_ids={4}, table_elements={5}, used_category_ids={6}, category_count={7}, elements_count={8}, errors={9}".format(
                ds.get("cat_id"),
                ds.get("schedule_id"),
                ds.get("used_table_data"),
                ds.get("table_rows"),
                ds.get("table_element_ids"),
                ds.get("table_elements"),
                ds.get("used_category_ids"),
                ds.get("category_count"),
                ds.get("elements_count"),
                ds.get("errors", []),
            )

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": "Заполнение завершено. Режим: {0}, Параметр листа: {1}".format(
                kit_mode, sheet_param_name
            ),
            "info": {
                "doc_title": doc.Title,
                "kit_mode": kit_mode,
                "sheet_param_name": sheet_param_name,
                "param_source": param_source,
                "all_browser_params": all_browser_params,
                "debug_info": debug_info,
                "debug_error": debug_error,
            },
            "fill": {
                "target_param": "ADSK_КомплектШифр",
                "source_param": sheet_param_name,
                "filled": fill_result["filled"],
                "planned_value": fill_result.get("planned_value"),
                "total": fill_result["total"],
                "updated_count": fill_result["updated_count"],
                "skipped_count": fill_result["skipped_count"],
                "skip_reasons": fill_result["skip_reasons"],
                "values": fill_result["values"],
                "message": fill_result["message"],
            },
        }

        if not fill_result["filled"] and fill_result["total"] == 0:
            ds = fill_result.get("debug_schedule")
            if ds:
                debug_msg = " [DEBUG: cat_id={0}, schedule_id={1}, used_collector_view={2}, used_cat_ids={3}, cat_count={4}, elem_count={5}, errors={6}]".format(
                    ds.get("cat_id"),
                    ds.get("schedule_id"),
                    ds.get("used_collector_view"),
                    ds.get("used_category_ids"),
                    ds.get("category_count"),
                    ds.get("elements_count"),
                    ds.get("errors", [])[:2],
                )
            else:
                debug_msg = ""
            result["fill"]["message"] = (
                "Заполнение не требовалось: нет элементов для обработки" + debug_msg
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
            "parameters": param_result["parameters"],
            "info": {
                "doc_title": doc.Title,
                "kit_mode": kit_mode,
                "sheet_param_name": sheet_param_name,
                "param_source": param_source,
                "all_browser_params": all_browser_params,
                "debug_info": debug_info,
                "debug_error": debug_error,
            },
            "fill": {
                "target_param": "ADSK_КомплектШифр",
                "source_param": sheet_param_name,
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
