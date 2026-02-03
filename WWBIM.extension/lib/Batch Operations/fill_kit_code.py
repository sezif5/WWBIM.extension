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
)
import System
from System.Collections.Generic import List

from add_shared_parameter import AddSharedParameterToDoc
from model_categories import MODEL_CATEGORIES


CONFIG = {
    "PARAMETER_NAME": "ADSK_КомплектШифр",
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


def DetermineKitMode(doc):
    filename = doc.Title.upper()
    prefixes = ["AR", "AI", "НАВ", "KM", "KR"]

    for prefix in prefixes:
        if filename.startswith(prefix):
            return "Schedules"

    return "Views"


def GetBrowserOrganizationSheetParameter(doc):
    try:
        from Autodesk.Revit.DB import BrowserOrganization

        org = BrowserOrganization.GetCurrentBrowserOrganizationForViews(doc)
        if org and org.Parameters.Size > 0:
            return org.Parameters.Item(0).GetName()
    except:
        pass
    return None


def DetermineSheetParameterName(doc):
    param_name = GetBrowserOrganizationSheetParameter(doc)
    if param_name:
        return param_name

    priority_params = [
        "ADSK_Штамп Раздел проекта",
        "ADSK_Раздел проекта",
        "Раздел проекта",
        "Sheet Number",
        "Номер листа",
    ]

    sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

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
            return param_name

    return None


def GetSheetParameter(sheet, param_name):
    param = sheet.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsString()
    return None


def GetCategoryFromSchedule(scheduleInstance):
    try:
        schedule_def = scheduleInstance.ScheduleDefinition
        return schedule_def.CategoryId
    except:
        return None


def GetElementsFromSchedule(doc, scheduleInstance):
    try:
        schedule_def = scheduleInstance.ScheduleDefinition
        elements = []

        cat_id = GetCategoryFromSchedule(scheduleInstance)
        if cat_id and cat_id.IntegerValue > 0:
            cat = doc.GetElement(cat_id)
            if cat:
                collector = FilteredElementCollector(doc)
                elements = (
                    collector.OfCategory(cat.Id.IntegerValue)
                    .WhereElementIsNotElementType()
                    .ToElements()
                )

        return elements
    except:
        return []


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
    param = element.LookupParameter("ADSK_КомплектШифр")
    if param and param.StorageType == StorageType.String:
        try:
            current_value = param.AsString()
            # Если уже есть значение, выбираем более короткий шифр
            if current_value:
                if len(kit_code) < len(current_value):
                    param.Set(kit_code)
                    return True
                else:
                    return False  # Текущий шифр короче или равен - не меняем
            else:
                # Нет значения - устанавливаем новый
                param.Set(kit_code)
                return True
        except:
            pass
    return False


def FillKitCodes_Schedules(doc, sheet_param_name, progress_callback=None):
    schedule_instances = (
        FilteredElementCollector(doc).OfClass(ScheduleSheetInstance).ToElements()
    )

    total_elements = 0
    updated_elements = 0
    skipped_elements = 0
    all_values = set()

    total_schedules = len(schedule_instances)
    current_schedule = 0

    for schedule_inst in schedule_instances:
        sheet_id = schedule_inst.OwnerSheetId
        if sheet_id.IntegerValue == -1:
            continue

        sheet = doc.GetElement(sheet_id)
        if not sheet:
            continue

        kit_code = GetSheetParameter(sheet, sheet_param_name)
        if not kit_code:
            continue

        all_values.add(kit_code)
        elements = GetElementsFromSchedule(doc, schedule_inst)
        total_elements += len(elements)

    current_schedule = 0

    for schedule_inst in schedule_instances:
        if progress_callback:
            progress = int((current_schedule / total_schedules) * 100)
            progress_callback(progress)

        sheet_id = schedule_inst.OwnerSheetId
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

        elements = GetElementsFromSchedule(doc, schedule_inst)

        for element in elements:
            current_value = GetKitCodeParameter(element)
            if not current_value:
                if SetKitCodeParameter(element, kit_code):
                    updated_elements += 1
                else:
                    skipped_elements += 1
            elif len(kit_code) < len(current_value):
                if SetKitCodeParameter(element, kit_code):
                    updated_elements += 1
                else:
                    skipped_elements += 1
            else:
                skipped_elements += 1

        current_schedule += 1

    filled = updated_elements > 0
    return {
        "total_elements": total_elements,
        "updated_elements": updated_elements,
        "skipped_elements": skipped_elements,
        "values": sorted(list(all_values)),
        "filled": filled,
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

    total_elements = 0
    updated_elements = 0
    skipped_elements = 0
    all_values = set()

    for sheet in sheets:
        kit_code = GetSheetParameter(sheet, sheet_param_name)
        if not kit_code:
            continue

        all_values.add(kit_code)
        view_ids = sheet.GetAllPlacedViews()

        for view_id in view_ids:
            view = doc.GetElement(view_id)
            if not view:
                continue

            if view.ViewType in view_types:
                elements = GetElementsFromView(doc, view)
                total_elements += len(elements)

    for sheet in sheets:
        kit_code = GetSheetParameter(sheet, sheet_param_name)
        if not kit_code:
            continue

        view_ids = sheet.GetAllPlacedViews()

        for view_id in view_ids:
            view = doc.GetElement(view_id)
            if not view:
                continue

            if view.ViewType in view_types:
                elements = GetElementsFromView(doc, view)

                for element in elements:
                    current_value = GetKitCodeParameter(element)
                    if not current_value:
                        if SetKitCodeParameter(element, kit_code):
                            updated_elements += 1
                        else:
                            skipped_elements += 1
                    elif len(kit_code) < len(current_value):
                        if SetKitCodeParameter(element, kit_code):
                            updated_elements += 1
                        else:
                            skipped_elements += 1
                    else:
                        skipped_elements += 1

    filled = updated_elements > 0
    return {
        "total_elements": total_elements,
        "updated_elements": updated_elements,
        "skipped_elements": skipped_elements,
        "values": sorted(list(all_values)),
        "filled": filled,
    }


def Execute(doc, progress_callback=None):
    kit_mode = DetermineKitMode(doc)
    sheet_param_name = DetermineSheetParameterName(doc)

    if not sheet_param_name:
        return {
            "success": False,
            "message": "Не удалось определить параметр листа для шифра комплекта",
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": "ADSK_КомплектШифр",
                "source_param": sheet_param_name,
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
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
        if kit_mode == "Schedules":
            fill_result = FillKitCodes_Schedules(
                doc, sheet_param_name, progress_callback
            )
        else:
            fill_result = FillKitCodes_Views(doc, sheet_param_name, progress_callback)

        t.Commit()

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": "Заполнение завершено. Режим: {0}, Параметр листа: {1}".format(
                kit_mode, sheet_param_name
            ),
            "fill": {
                "target_param": "ADSK_КомплектШифр",
                "source_param": sheet_param_name,
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
            "parameters": param_result["parameters"],
            "fill": {
                "target_param": "ADSK_КомплектШифр",
                "source_param": sheet_param_name,
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
