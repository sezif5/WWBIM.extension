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
    BuiltInCategory,
    StorageType,
    Transaction,
)


CONFIG = {
    "PARAM_1": "ДГП_Имя уровня",
    "PARAM_2": "Имя",
}


def _to_str(value):
    if value is None:
        return ""
    try:
        return value
    except Exception:
        try:
            return str(value)
        except Exception:
            return ""


def _normalize(value):
    return _to_str(value).replace("\u00a0", " ").strip()


def _get_levels(doc):
    return (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Levels)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def _get_param_text(level, param_name):
    try:
        p = level.LookupParameter(param_name)
        if not p or not p.HasValue:
            return None
        if p.StorageType == StorageType.String:
            return p.AsString()
        vs = p.AsValueString()
        return vs if vs else None
    except Exception:
        return None


def _set_param_text(level, param_name, value):
    try:
        p = level.LookupParameter(param_name)
        if not p:
            return {"status": "parameter_not_found", "reason": "parameter_not_found"}
        if p.IsReadOnly:
            return {"status": "readonly", "reason": "readonly"}
        if p.StorageType != StorageType.String:
            return {"status": "wrong_storage_type", "reason": "wrong_storage_type"}

        current = p.AsString()
        if _normalize(current) == _normalize(value):
            return {"status": "already_ok", "reason": "already_ok"}

        p.Set(_to_str(value))
        return {"status": "updated", "reason": None}
    except Exception:
        return {"status": "exception", "reason": "exception"}


def _set_level_name(level, value):
    try:
        current = _to_str(level.Name)
        if _normalize(current) == _normalize(value):
            return {"status": "already_ok", "reason": "already_ok"}

        level.Name = _to_str(value)
        return {"status": "updated", "reason": None}
    except Exception:
        # fallback: попробовать через параметр
        return _set_param_text(level, CONFIG["PARAM_2"], value)


def SwapLevelNames(doc, progress_callback=None):
    levels = _get_levels(doc)
    total = len(levels)

    skip_reasons = {
        "both_empty": 0,
        "parameter_not_found": 0,
        "readonly": 0,
        "wrong_storage_type": 0,
        "already_ok": 0,
        "exception": 0,
    }

    updated_count = 0
    skipped_count = 0
    swapped_values = []

    if total == 0:
        return {
            "total": 0,
            "updated_count": 0,
            "skipped_count": 0,
            "skip_reasons": skip_reasons,
            "filled": False,
            "values": [],
            "message": "Нет уровней для обработки",
        }

    for idx, level in enumerate(levels):
        if progress_callback:
            progress = int(((idx + 1) / float(total)) * 100)
            progress_callback(progress)

        val1_old = _get_param_text(level, CONFIG["PARAM_1"])
        val2_old = _to_str(getattr(level, "Name", None))

        if (not _normalize(val1_old)) and (not _normalize(val2_old)):
            skipped_count += 1
            skip_reasons["both_empty"] += 1
            continue

        # обмен: параметр1 <- старое имя, имя <- старый параметр1
        target_param1 = val2_old
        target_name = val1_old

        res_p1 = _set_param_text(level, CONFIG["PARAM_1"], target_param1)
        res_name = _set_level_name(level, target_name)

        # успешным считаем, если хотя бы одно из двух реально обновилось
        changed = (res_p1.get("status") == "updated") or (
            res_name.get("status") == "updated"
        )

        if changed:
            updated_count += 1
            swapped_values.append(_to_str(target_name))
        else:
            skipped_count += 1
            reason = None
            if res_name.get("reason") not in (None, "already_ok"):
                reason = res_name.get("reason")
            elif res_p1.get("reason") not in (None, "already_ok"):
                reason = res_p1.get("reason")
            else:
                reason = "already_ok"
            if reason in skip_reasons:
                skip_reasons[reason] += 1
            else:
                skip_reasons["exception"] += 1

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
        "filled": filled,
        "values": sorted(list(set(swapped_values))),
        "message": message,
    }


def Execute(doc, progress_callback=None):
    t = None

    try:
        if not doc.IsModifiable:
            t = Transaction(doc, "Обмен значений ДГП_Имя уровня и Имя")
            t.Start()

        fill_result = SwapLevelNames(doc, progress_callback)

        if t is not None:
            t.Commit()

        result = {
            "success": True,
            "message": "Обмен значений уровней завершен",
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": "{0} <-> {1}".format(
                    CONFIG["PARAM_1"], CONFIG["PARAM_2"]
                ),
                "source": "Взаимный обмен значений",
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
            result["fill"]["message"] = "Обмен не требовался: уровни не найдены"
        elif not fill_result["filled"]:
            result["fill"]["message"] = "Обмен не требовался: обновлений нет"

        return result

    except Exception as e:
        if t is not None:
            try:
                t.RollBack()
            except Exception:
                pass

        return {
            "success": False,
            "message": "Ошибка: {0}".format(str(e)),
            "parameters": {"added": [], "existing": [], "failed": []},
            "fill": {
                "target_param": "{0} <-> {1}".format(
                    CONFIG["PARAM_1"], CONFIG["PARAM_2"]
                ),
                "source": "Взаимный обмен значений",
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
