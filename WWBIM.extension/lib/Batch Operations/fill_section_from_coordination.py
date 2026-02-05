# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import sys
import os

from Autodesk.Revit.DB import (
    BuiltInParameter,
    BuiltInCategory,
    Category,
    ElementId,
    FilteredElementCollector,
    Transaction,
    TransactionStatus,
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

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, ".."))
sys.path.insert(0, LIB_DIR)


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

import openbg
import closebg
from model_categories import MODEL_CATEGORIES
from add_shared_parameter import AddSharedParameterToDoc

# ---------- Config ----------

CONFIG = {
    "PARAMETER_NAME": "ADSK_Секция",
    "BINDING_TYPE": "instance",
    "PARAMETER_GROUP": "PG_IDENTITY_DATA",
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


def FindCoordinationFile(doc, objects_dir):
    """Найти координационный txt файл и получить путь к модели с _CR_."""
    model_name = os.path.splitext(doc.Title)[0]

    # Формируем имя координационного txt файла (добавляем _SERV)
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
    if not os.path.isdir(OBJECTS_DIR):
        return {
            "success": False,
            "message": "Папка Objects не найдена: {0}".format(OBJECTS_DIR),
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": False,
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }

    coord_path, error = FindCoordinationFile(doc, OBJECTS_DIR)
    if error:
        return {
            "success": False,
            "message": error,
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": False,
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }

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
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
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
                "planned_value": None,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
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
                "fill": {
                    "target_param": SECTION_PARAM,
                    "source": "Координационный файл",
                    "filled": False,
                    "planned_value": None,
                    "total": 0,
                    "updated_count": 0,
                    "skipped_count": 0,
                    "skip_reasons": {},
                    "values": [],
                },
            }

        # Подготавливаем фильтр
        mcat_filter = multicategory_filter()
        assigned = {}
        total = 0
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

        # Проходим по объёмам и заполняем секции
        for i, v in enumerate(volumes):
            solid = v["solid"]
            section = v["section"]
            all_values.add(section)

            if progress_callback:
                progress = int((i / float(len(volumes))) * 100)
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

                total += 1

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

                ok_any = False
                for tgt in iter_with_subcomponents(el):
                    result = SetParameterValue(tgt, SECTION_PARAM, section)
                    if result["status"] == "updated":
                        ok_any = True
                    else:
                        reason = result["reason"]
                        if reason in skip_reasons:
                            skip_reasons[reason] += 1
                        fails.append([str(tgt.Id), family_label(tgt), reason])

                if ok_any:
                    updated_count += 1
                    assigned[eid] = section
                else:
                    skipped_count += 1

        filled = updated_count > 0

        # Формируем сообщение с проблемами
        info_messages = []
        if skipped:
            info_messages.append("Пропущено объёмов: {0}".format(len(skipped)))
        if conflicts:
            info_messages.append("Конфликтов: {0}".format(len(conflicts)))
        if fails:
            info_messages.append("Ошибок записи: {0}".format(len(fails)))

        reasons_str = "; ".join(
            ["{0}={1}".format(k, v) for k, v in skip_reasons.items() if v > 0]
        )
        message_parts = []
        message_parts.append("total={0}".format(total))
        message_parts.append("updated={0}".format(updated_count))
        message_parts.append("skipped={0}".format(skipped_count))
        if reasons_str:
            message_parts.append("reasons: " + reasons_str)
        if info_messages:
            message_parts.append("; ".join(info_messages))

        message = ", ".join(message_parts)

        return {
            "success": True,
            "message": message,
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
                "filled": filled,
                "planned_value": sorted(list(all_values)),
                "total": total,
                "updated_count": updated_count,
                "skipped_count": skipped_count,
                "skip_reasons": skip_reasons,
                "values": sorted(list(all_values)),
                "message": message,
            },
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
    param_result = {
        "success": False,
        "parameters": {"added": [], "existing": [], "failed": []},
        "message": "Параметр не был добавлен",
    }

    t = None

    try:
        # Если документ уже модифицируется - работаем без транзакции
        if doc.IsModifiable:
            param_result = EnsureParameterExists(doc)
            if not param_result["success"]:
                return {
                    "success": False,
                    "message": param_result["message"],
                    "parameters": param_result["parameters"],
                }

            fill_result = FillSectionsFromCoordination(doc, progress_callback)

            result = {
                "success": fill_result["success"],
                "parameters": param_result["parameters"],
                "message": fill_result["message"],
                "fill": fill_result["fill"],
            }

            return result

        # Иначе стартуем свою транзакцию
        t = Transaction(doc, "Заполнение параметра ADSK_Номер секции")
        t.Start()

        param_result = EnsureParameterExists(doc)
        if not param_result["success"]:
            try:
                t.RollBack()
            except Exception:
                pass
            return {
                "success": False,
                "message": param_result["message"],
                "parameters": param_result["parameters"],
            }

        fill_result = FillSectionsFromCoordination(doc, progress_callback)

        t.Commit()

        result = {
            "success": fill_result["success"],
            "parameters": param_result["parameters"],
            "message": fill_result["message"],
            "fill": fill_result["fill"],
        }

        return result

    except Exception as e:
        # Откатываем только если транзакция была стартована
        if t and t.GetStatus() == TransactionStatus.Started:
            try:
                t.RollBack()
            except Exception:
                pass

        return {
            "success": False,
            "message": "Ошибка: {0}".format(str(e)),
            "parameters": param_result["parameters"],
            "fill": {
                "target_param": SECTION_PARAM,
                "source": "Координационный файл",
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
