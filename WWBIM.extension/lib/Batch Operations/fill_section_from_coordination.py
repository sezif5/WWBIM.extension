# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import sys
import os
import inspect

from Autodesk.Revit.DB import (
    BuiltInParameter,
    BuiltInCategory,
    Category,
    ElementId,
    FilteredElementCollector,
    Transaction,
    StorageType,
    FamilyInstance,
    RevitLinkInstance,
    Solid,
    SolidUtils,
    ElementIntersectsSolidFilter,
    ElementMulticategoryFilter,
    BooleanOperationsUtils,
    BooleanOperationsType,
    Options,
    ViewDetailLevel,
    BuiltInParameterGroup,
)

from System.Collections.Generic import List

SCRIPT_DIR = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
LIB_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, ".."))
sys.path.insert(0, LIB_DIR)

import openbg
import closebg
from model_categories import MODEL_CATEGORIES
from add_shared_parameter import (
    AddSharedParameterToDoc,
    BindParameter,
    GetSharedParameterFile,
)

out = script.get_output()

# ---------- Config ----------

CONFIG = {
    "PARAMETER_NAME": "ADSK_Секция",
    "BINDING_TYPE": "instance",
    "PARAMETER_GROUP": None,
    "CATEGORIES": MODEL_CATEGORIES,
}

SECTION_PARAM = "ADSK_Секция"
VOL_CAT = BuiltInCategory.OST_Entourage

TARGET_CATS = [
    # ОГС
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Columns,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_StructuralStiffener,
    BuiltInCategory.OST_Stairs,
    BuiltInCategory.OST_Railings,
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_CurtainWallMullions,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_GenericModel,
    # MEP
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_Sprinklers,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_CableTrayFitting,
    BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_ConduitFitting,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_LightingDevices,
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_SpecialityEquipment,
]

# ---------- Helpers ----------


def GetParameterValue(element, param_name):
    param = element.LookupParameter(param_name)
    if param and param.HasValue:
        try:
            if param.StorageType == StorageType.String:
                return param.AsString()
            elif param.StorageType == StorageType.ElementId:
                elem_id = param.AsElementId()
                if elem_id and elem_id.IntegerValue != -1:
                    elem = element.Document.GetElement(elem_id)
                    return elem.Name if elem else None
            elif param.StorageType == StorageType.Integer:
                return str(param.AsInteger())
            elif param.StorageType == StorageType.Double:
                return param.AsValueString()
        except Exception:
            pass
    return None


def SetParameterValue(element, param_name, value):
    param = element.LookupParameter(param_name)
    if not param or param.IsReadOnly:
        return False
    try:
        param.Set(value)
        return True
    except Exception:
        return False


def solids_of_element(el):
    """Вернёт единый Solid элемента (объединение), либо пустой список."""
    try:
        opt = Options()
        opt.DetailLevel = ViewDetailLevel.Fine
        opt.IncludeNonVisibleObjects = True
        geo = el.get_Geometry(opt)
        if not geo:
            return []

        def _acc(giter, cur):
            for g in giter:
                if isinstance(g, Solid) and g.Volume > 1e-9:
                    cur = (
                        g
                        if cur is None
                        else BooleanOperationsUtils.ExecuteBooleanOperation(
                            cur, g, BooleanOperationsType.Union
                        )
                    )
                elif hasattr(g, "GetInstanceGeometry"):
                    cur = _acc(g.GetInstanceGeometry(), cur)
            return cur

        union = _acc(geo, None)
        return [union] if union else []
    except Exception:
        return []


def multicategory_filter():
    ids = List[ElementId]()
    for bic in TARGET_CATS:
        ids.Add(ElementId(int(bic)))
    return ElementMulticategoryFilter(ids)


def family_label(el):
    try:
        et = el.Document.GetElement(el.GetTypeId())
        if et:
            fam = getattr(et, "FamilyName", None)
            typ = getattr(et, "Name", None)
            if fam and typ:
                return "%s : %s" % (fam, typ)
            if fam:
                return fam
            if typ:
                return typ
    except Exception:
        pass
    try:
        if hasattr(el, "Symbol") and el.Symbol:
            return "%s : %s" % (el.Symbol.Family.Name, el.Symbol.Name)
    except Exception:
        pass
    return el.Category.Name if el.Category else ""


def iter_with_subcomponents(root):
    """Сам элемент + все вложенные FamilyInstance подкомпоненты (без дублей)."""
    stack = [root]
    visited = set([root.Id.IntegerValue])
    while stack:
        el = stack.pop()
        yield el
        if isinstance(el, FamilyInstance):
            try:
                for sid in el.GetSubComponentIds() or []:
                    if sid.IntegerValue in visited:
                        continue
                    sub = root.Document.GetElement(sid)
                    if sub:
                        visited.add(sid.IntegerValue)
                        stack.append(sub)
            except Exception:
                pass


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


def FindCoordinationFile(doc, objects_dir):
    """Найти txt файл объекта и определить путь к координационному файлу."""
    model_name = os.path.splitext(doc.Title)[0]

    # Ищем txt файл с именем модели
    txt_path = os.path.join(objects_dir, model_name + ".txt")
    if not os.path.exists(txt_path):
        return None, "Не найден txt файл объекта: {0}".format(model_name + ".txt")

    # Формируем путь к координационному txt файлу (добавляем _SERV)
    coord_txt_name = model_name + "_SERV.txt"
    coord_txt_path = os.path.join(objects_dir, coord_txt_name)
    if not os.path.exists(coord_txt_path):
        return None, "Не найден координационный txt файл: {0}".format(coord_txt_name)

    # Читаем путь к модели из координационного txt файла
    try:
        with open(coord_txt_path, "rb") as f:
            for raw in f:
                try:
                    model_path = raw.decode("utf-8").strip()
                except Exception:
                    try:
                        model_path = raw.decode("cp1251").strip()
                    except Exception:
                        model_path = raw.strip()
                if model_path:
                    model_path = str(model_path)
                    # Проверяем что путь содержит _CR_
                    if "_CR_" in model_path:
                        return model_path, None
                    else:
                        return (
                            None,
                            "В координационном txt файле указан путь без _CR_: {0}".format(
                                model_path
                            ),
                        )
    except Exception as e:
        return None, "Ошибка чтения координационного txt файла: {0}".format(str(e))

    return None, "Координационный txt файл пуст или содержит ошибки"


# ---------- Main Logic ----------


def FillSectionsFromCoordination(doc, progress_callback=None):
    volumes = []
    skipped = []
    conflicts = []
    fails = []

    # Находим путь к координационному файлу
    objects_dir = os.path.join(os.path.dirname(doc.PathName), "Objects")
    if not os.path.isdir(objects_dir):
        return {
            "success": False,
            "message": "Папка Objects не найдена",
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }

    coord_path, error = FindCoordinationFile(doc, objects_dir)
    if error:
        return {
            "success": False,
            "message": error,
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }

    out.print_md("Координационный файл: `{0}`".format(coord_path))

    # Открываем координационный файл в фоне
    coord_doc = None
    try:
        result = openbg.open_in_background(
            __revit__.Application,
            __revit__,
            coord_path,
            audit=False,
            worksets="all",
            detach=False,
            suppress_warnings=True,
        )
        if result and len(result) >= 1:
            coord_doc = result[0]
        else:
            raise Exception("openbg не вернул документ")
    except Exception as e:
        return {
            "success": False,
            "message": "Ошибка открытия координационного файла: {0}".format(str(e)),
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }

    if not coord_doc:
        return {
            "success": False,
            "message": "Координационный документ не загружен",
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }

    try:
        # Находим объёмы в координационном файле
        vol_collector = (
            FilteredElementCollector(coord_doc)
            .OfCategory(VOL_CAT)
            .WhereElementIsNotElementType()
        )

        for vol_elem in vol_collector:
            sols = solids_of_element(vol_elem)
            if not sols:
                skipped.append("{0} (нет геометрии)".format(family_label(vol_elem)))
                continue

            section = GetParameterValue(vol_elem, SECTION_PARAM)
            if not section:
                skipped.append("{0} (нет номера секции)".format(family_label(vol_elem)))
                continue

            volumes.append(
                {"solid": sols[0], "section": section, "label": family_label(vol_elem)}
            )

        if not volumes:
            return {
                "success": False,
                "message": "Не найдены объёмы (Антураж) с номером секции",
                "parameters": {"added": [], "existing": [], "failed": []},
                "fill": {
                    "target_param": SECTION_PARAM,
                    "source": "Координационный файл",
                    "filled": False,
                    "total_elements": 0,
                    "updated_elements": 0,
                    "skipped_elements": 0,
                    "values": [],
                },
            }

        out.print_md(
            "Найдено объёмов (Антураж): {0}, секций: {1}".format(
                len(volumes), len(set(v["section"] for v in volumes))
            )
        )

        # Подготавливаем фильтр
        mcat_filter = multicategory_filter()
        assigned = {}
        total_elements = 0
        updated_elements = 0
        skipped_elements = 0
        all_values = set()

        # Проходим по объёмам и заполняем секции
        for v in volumes:
            solid = v["solid"]
            section = v["section"]
            all_values.add(section)

            if progress_callback:
                progress = int((volumes.index(v) / float(len(volumes))) * 100)
                progress_callback(progress)

            col = (
                FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WherePasses(mcat_filter)
                .WherePasses(ElementIntersectsSolidFilter(solid))
            )

            for el in col:
                # Пропускаем экземпляры связей
                if isinstance(el, RevitLinkInstance):
                    continue

                total_elements += 1

                eid = el.Id.IntegerValue

                # Если элемент уже попал в другой объём с другой секцией - конфликт
                if eid in assigned and assigned[eid] != section:
                    conflicts.append(
                        [
                            str(el.Id),
                            family_label(el),
                            "{0} -> {1}".format(assigned[eid], section),
                        ]
                    )
                    continue

                ok_any = True
                for tgt in iter_with_subcomponents(el):
                    ok, reason = SetParameterValue(tgt, SECTION_PARAM, section)
                    if not ok:
                        ok_any = False
                        fails.append([str(tgt.Id), family_label(tgt), reason or ""])

                if ok_any:
                    updated_elements += 1
                    assigned[eid] = section
                else:
                    skipped_elements += 1

        filled = updated_elements > 0

        return {
            "total_elements": total_elements,
            "updated_elements": updated_elements,
            "skipped_elements": skipped_elements,
            "values": sorted(list(all_values)),
            "conflicts": conflicts,
            "fails": fails,
            "skipped_volumes": skipped,
            "filled": filled,
        }

    finally:
        # Закрываем координационный файл
        try:
            closebg.close_with_policy(
                coord_doc, do_sync=False, comment="Заполнение секций"
            )
        except Exception:
            try:
                coord_doc.Close(False)
            except Exception:
                pass


# ---------- Execute ----------


def Execute(doc, progress_callback=None):
    t = Transaction(doc, "Заполнение параметра ADSK_Номер секции")
    t.Start()

    param_result = {
        "success": False,
        "parameters": {"added": [], "existing": [], "failed": []},
        "message": "Параметр не был добавлен",
    }

    try:
        param_result = EnsureParameterExists(doc)
        fill_result = FillSectionsFromCoordination(doc, progress_callback)

        t.Commit()

        result = {
            "success": True,
            "parameters": param_result["parameters"],
            "message": param_result["message"],
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": fill_result["filled"],
                "total_elements": fill_result["total_elements"],
                "updated_elements": fill_result["updated_elements"],
                "skipped_elements": fill_result["skipped_elements"],
                "values": fill_result["values"],
            },
        }

        # Отчёты о проблемах
        if fill_result.get("skipped_volumes"):
            out.print_md("### Пропущенные объёмы")
            for msg in fill_result["skipped_volumes"]:
                out.print_md("- {0}".format(msg))

        if fill_result.get("conflicts"):
            out.print_md("### Конфликты (элемент попал в разные объёмы)")
            for c in fill_result["conflicts"]:
                out.print_md("- ID: {0}, {1}: {2}".format(c[0], c[1], c[2]))

        if fill_result.get("fails"):
            out.print_md("### Не удалось записать параметр")
            for f in fill_result["fails"]:
                out.print_md("- ID: {0}, {1}: {2}".format(f[0], f[1], f[2]))

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
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": False,
                "total_elements": 0,
                "updated_elements": 0,
                "skipped_elements": 0,
                "values": [],
            },
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
