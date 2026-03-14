# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import os
import re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, ".."))
sys.path.insert(0, LIB_DIR)

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    RevitLinkType,
    ElementId,
    BuiltInParameter,
    Transaction,
    Workset,
    WorksetKind,
    FilteredWorksetCollector,
    WorksetDefaultVisibilitySettings,
    WorksetVisibility,
    View,
    ModelPathUtils,
    ImportInstance,
    StorageType,
)


CONFIG = {
    "ALWAYS_PIN_AFTER_ASSIGN": True,
}


RULES = [
    ({"KOORD"}, "05_Links_KOORD", True),
    ({"AR"}, "05_Links_AR", False),
    ({"KR"}, "05_Links_KR", False),
    ({"OV"}, "05_Links_OV", False),
    ({"VK"}, "05_Links_VK", False),
    ({"VNS"}, "05_Links_VNS", False),
    ({"EOM"}, "05_Links_EOM", False),
    ({"SS"}, "05_Links_SS", False),
    ({"ITP"}, "05_Links_ITP", False),
    ({"PT"}, "05_Links_PT", False),
]

DWG_WS_NAME = "05_Links_DWG"

SPECIAL_MODELS = {"AI", "AR", "KM", "KR", "KG", "CR"}

LAT2CYR = {
    "A": "А",
    "B": "В",
    "C": "С",
    "E": "Е",
    "H": "Н",
    "K": "К",
    "M": "М",
    "O": "О",
    "P": "Р",
    "T": "Т",
    "X": "Х",
    "Y": "У",
}
SEG_SPLIT = re.compile(r"[\W_]+", re.UNICODE)


def normalize_cyr(s):
    if not s:
        return ""
    up = s.upper()
    return "".join(LAT2CYR.get(ch, ch) for ch in up)


def split_segments(text):
    norm = normalize_cyr(text or "")
    return [seg for seg in SEG_SPLIT.split(norm) if seg]


def segments_from_link(doc, ltype, linstr):
    segs = []
    try:
        segs += split_segments(getattr(ltype, "Name", ""))
    except Exception:
        pass
    try:
        segs += split_segments(getattr(linstr, "Name", ""))
    except Exception:
        pass
    try:
        efr = ltype.GetExternalFileReference()
        if efr:
            upath = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                efr.GetAbsolutePath()
            )
            if upath:
                base = os.path.splitext(os.path.basename(upath))[0]
                segs += split_segments(base)
    except Exception:
        pass
    seen, uniq = set(), []
    for s in segs:
        if s not in seen:
            uniq.append(s)
            seen.add(s)
    return uniq


def match_rule(segments):
    for tokens, wsname, hide in RULES:
        for tok in tokens:
            tok_norm = normalize_cyr(tok)
            for seg in segments:
                if seg == tok_norm or seg.startswith(tok_norm):
                    return wsname, hide, tok
    return None, None, None


def ensure_workset(doc, ws_name):
    existing = [
        ws
        for ws in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
        if ws.Name == ws_name
    ]
    if existing:
        return existing[0], False
    t = Transaction(doc, "Создать рабочий набор: {0}".format(ws_name))
    t.Start()
    ws = Workset.Create(doc, ws_name)
    t.Commit()
    return ws, True


def set_default_visibility(doc, workset, visible):
    settings = WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(doc)
    t = Transaction(doc, "Дефолтная видимость РН: {0}".format(workset.Name))
    t.Start()
    settings.SetVisibility(workset.Id, bool(visible))
    t.Commit()


def hide_workset_in_all_views(doc, workset):
    views = FilteredElementCollector(doc).OfClass(View)
    cnt = 0
    t = Transaction(doc, "Скрыть РН во всех видах: {0}".format(workset.Name))
    t.Start()
    for v in views:
        try:
            if v and not v.IsTemplate:
                v.SetWorksetVisibility(workset.Id, WorksetVisibility.Hidden)
                cnt += 1
        except Exception:
            pass
    t.Commit()
    return cnt


def assign_to_workset(elem, workset, stats):
    p = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
    if p is None:
        stats["no_param"] += 1
        return False, "нет параметра ELEM_PARTITION_PARAM"
    if p.IsReadOnly:
        stats["readonly"] += 1
        return False, "параметр только для чтения"
    try:
        if p.StorageType == StorageType.ElementId:
            ok = p.Set(workset.Id)
        else:
            ok = p.Set(workset.Id.IntegerValue)
        if not ok:
            stats["failed_set"] += 1
            return False, "Set(...) вернул False"
        if CONFIG["ALWAYS_PIN_AFTER_ASSIGN"] and (
            isinstance(elem, RevitLinkInstance) or isinstance(elem, ImportInstance)
        ):
            try:
                elem.Pinned = True
            except Exception:
                pass
        return True, "ok"
    except Exception as ex:
        stats["exceptions"] += 1
        return False, "исключение: {0}".format(ex)


def get_link_filename(ltype):
    try:
        efr = ltype.GetExternalFileReference()
        if efr:
            upath = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                efr.GetAbsolutePath()
            )
            if upath:
                base = os.path.splitext(os.path.basename(upath))[0]
                return base
    except Exception:
        pass
    return None


def model_has_special_token(doc):
    model_name = doc.Title
    if model_name.endswith("_отсоединено"):
        model_name = model_name[: -len("_отсоединено")]

    # Упрощенная логика:
    # если в имени модели есть один из тегов (как отдельный сегмент) -> режим по файлам,
    # иначе сразу режим по разделам.
    upper_name = model_name.upper()
    for token in SPECIAL_MODELS:
        pattern = r"(^|[\W_]){}($|[\W_])".format(re.escape(token))
        if re.search(pattern, upper_name):
            return True
    return False


def Execute(doc, progress_callback=None):
    if doc is None:
        return {
            "success": False,
            "message": "Документ не передан",
            "fill": {
                "target_param": "Worksets",
                "source": "Links",
                "filled": False,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }
    if doc.IsFamilyDocument:
        return {
            "success": False,
            "message": "Скрипт работает только в проектных файлах, а не в семействах.",
            "fill": {
                "target_param": "Worksets",
                "source": "Links",
                "filled": False,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }
    if not doc.IsWorkshared:
        return {
            "success": False,
            "message": "В модели не включено совместное использование. Включите «Рабочие наборы» и запустите скрипт снова.",
            "fill": {
                "target_param": "Worksets",
                "source": "Links",
                "filled": False,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }

    types = list(FilteredElementCollector(doc).OfClass(RevitLinkType))
    insts = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
    if not types or not insts:
        return {
            "success": False,
            "message": "В файле нет Revit-связей (типов или экземпляров).",
            "fill": {
                "target_param": "Worksets",
                "source": "Links",
                "filled": False,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }

    insts_by_type = {}
    for inst in insts:
        try:
            insts_by_type.setdefault(inst.GetTypeId(), []).append(inst)
        except Exception:
            pass

    plan = []
    required_ws = set()
    ws_to_hide = set()

    is_special_model = model_has_special_token(doc)

    for ltype in types:
        tname = getattr(ltype, "Name", "") or ""
        for inst in insts_by_type.get(ltype.Id, []):
            if is_special_model:
                link_filename = get_link_filename(ltype)
                if link_filename:
                    wsname = "05_Links_{0}".format(link_filename)
                    plan.append((inst, wsname, False))
                    required_ws.add(wsname)
            else:
                segs = segments_from_link(doc, ltype, inst)
                wsname, hide, tok = match_rule(segs)
                if wsname:
                    plan.append((inst, wsname, hide))
                    required_ws.add(wsname)
                    if hide:
                        ws_to_hide.add(wsname)

    cad_insts = list(FilteredElementCollector(doc).OfClass(ImportInstance))
    if cad_insts:
        for cad in cad_insts:
            plan.append((cad, DWG_WS_NAME, False))
        required_ws.add(DWG_WS_NAME)

    if not plan:
        if is_special_model:
            message = "В файле нет RVT-связей."
        else:
            message = "Подходящих RVT-связей по заданным токенам не найдено."
        return {
            "success": False,
            "message": message,
            "fill": {
                "target_param": "Worksets",
                "source": "Links",
                "filled": False,
                "total": 0,
                "updated_count": 0,
                "skipped_count": 0,
                "skip_reasons": {},
                "values": [],
            },
        }

    ws_cache, created_ws = {}, []
    for name in sorted(required_ws):
        ws, created = ensure_workset(doc, name)
        ws_cache[name] = ws
        if created:
            created_ws.append(name)

    stats = {"no_param": 0, "readonly": 0, "failed_set": 0, "exceptions": 0}
    moved = 0
    details = []
    skip_reasons = {
        "no_param": 0,
        "readonly": 0,
        "failed_set": 0,
        "exceptions": 0,
    }

    if progress_callback:
        progress_callback(0)

    total_items = len(plan)
    current_index = 0

    t = Transaction(doc, "Распределить связи по рабочим наборам")
    t.Start()
    for inst, wsname, _hide in plan:
        if progress_callback:
            progress = int(((current_index + 1) / float(total_items)) * 100)
            progress_callback(progress)
        ok, msg = assign_to_workset(inst, ws_cache[wsname], stats)
        if ok:
            moved += 1
        else:
            if "нет параметра" in msg:
                skip_reasons["no_param"] += 1
            elif "только для чтения" in msg:
                skip_reasons["readonly"] += 1
            elif "Set(...) вернул False" in msg:
                skip_reasons["failed_set"] += 1
            else:
                skip_reasons["exceptions"] += 1
        details.append(
            [getattr(inst, "Name", ""), wsname, "OK" if ok else ("FAIL: " + msg)]
        )
        current_index += 1
    t.Commit()

    hidden_report = []
    for name in sorted(ws_to_hide):
        ws = ws_cache.get(name)
        if not ws:
            continue
        try:
            set_default_visibility(doc, ws, False)
        except Exception as ex:
            pass
        try:
            hidden_cnt = hide_workset_in_all_views(doc, ws)
        except Exception as ex:
            hidden_cnt = 0
        hidden_report.append((name, hidden_cnt))

    all_values = sorted(list(required_ws))
    filled = moved > 0

    reasons_str = "; ".join(
        ["{0}={1}".format(k, v) for k, v in skip_reasons.items() if v > 0]
    )
    message_parts = []
    message_parts.append("moved={0}".format(moved))
    message_parts.append("total={0}".format(total_items))
    if created_ws:
        message_parts.append("created_ws={0}".format(len(created_ws)))
    if reasons_str:
        message_parts.append("skip_reasons: {0}".format(reasons_str))
    if hidden_report:
        message_parts.append("hidden_in_views={0}".format(len(hidden_report)))

    message = ", ".join(message_parts)

    return {
        "success": True,
        "message": message,
        "fill": {
            "target_param": "Worksets",
            "source": "Links",
            "filled": filled,
            "total": total_items,
            "updated_count": moved,
            "skipped_count": total_items - moved,
            "skip_reasons": skip_reasons,
            "values": all_values,
            "message": message,
            "created_ws": created_ws,
            "hidden_report": hidden_report,
            "details": details,
        },
    }


if __name__ == "__main__":
    import traceback

    try:
        Execute(__revit__.ActiveUIDocument.Document)
    except Exception as e:
        print("Ошибка: {0}\n{1}".format(e, traceback.format_exc()))
