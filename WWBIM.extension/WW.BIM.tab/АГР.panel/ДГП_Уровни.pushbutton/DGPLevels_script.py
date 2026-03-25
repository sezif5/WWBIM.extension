# -*- coding: utf-8 -*-
"""
ДГП Уровни — копирование названия уровня из параметра ДГП_Имя уровня в параметр Имя.
"""

from __future__ import print_function, division

from pyrevit import revit, DB, forms, script


doc = revit.doc
output = script.get_output()


# ============================================================================
# PARAM NAMES
# ============================================================================

PARAM_SOURCE_NAME = "ДГП_Имя уровня"
PARAM_TARGET_NAME = "Имя"


# ============================================================================
# HELPERS
# ============================================================================


def to_unicode(v):
    if v is None:
        return ""
    try:
        return unicode(v)
    except Exception:
        try:
            return str(v)
        except Exception:
            return ""


def normalize_for_compare(v):
    s = to_unicode(v).replace("\u00a0", " ").strip()
    return s


def get_param(elem, name):
    if not elem or not name:
        return None
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None


def get_param_text(param):
    if not param or not param.HasValue:
        return None
    try:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
        vs = param.AsValueString()
        if vs:
            return vs
    except Exception:
        pass
    return None


def set_param_value(elem, param_name, value):
    param = get_param(elem, param_name)
    if not param:
        return {
            "status": "not_found",
            "reason": "parameter_not_found",
            "param": param_name,
        }
    if param.IsReadOnly:
        return {
            "status": "readonly",
            "reason": "readonly",
            "param": param_name,
        }
    if param.StorageType != DB.StorageType.String:
        return {
            "status": "wrong_type",
            "reason": "wrong_storage_type",
            "param": param_name,
        }
    current = get_param_text(param)
    if normalize_for_compare(current) == normalize_for_compare(value):
        return {
            "status": "already_ok",
            "reason": "already_ok",
            "param": param_name,
        }
    try:
        param.Set(to_unicode(value))
        return {
            "status": "updated",
            "reason": None,
            "param": param_name,
        }
    except Exception as e:
        return {
            "status": "exception",
            "reason": to_unicode(e),
            "param": param_name,
        }


# ============================================================================
# MAIN PROCESSING
# ============================================================================


def process_levels():
    stats = {
        "total_processed": 0,
        "total_updated": 0,
        "total_skipped": 0,
        "skip_reasons": {
            "source_empty": 0,
            "parameter_not_found": 0,
            "readonly": 0,
            "already_ok": 0,
            "wrong_storage_type": 0,
            "exception": 0,
        },
        "results": [],
    }

    try:
        levels = (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_Levels)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception as e:
        print("Ошибка при сборе уровней: {}".format(to_unicode(e)))
        output.print_md("### Ошибка при сборе уровней")
        output.print_md("```\n{}\n```".format(to_unicode(e)))
        return stats

    skip_labels = {
        "source_empty": "Пустое значение в ДГП_Имя уровня",
        "parameter_not_found": "Не найден целевой параметр",
        "readonly": "Параметр только для чтения",
        "already_ok": "Значение уже совпадает",
        "wrong_storage_type": "Неверный тип параметра",
        "exception": "Ошибка при записи",
    }

    for level in levels:
        stats["total_processed"] += 1

        try:
            level_id = level.Id.IntegerValue
        except Exception:
            level_id = -1

        source_param = get_param(level, PARAM_SOURCE_NAME)
        source_value = get_param_text(source_param)

        if not source_value or source_value.strip() == "":
            stats["skip_reasons"]["source_empty"] += 1
            stats["total_skipped"] += 1
            stats["results"].append(
                {
                    "id": level_id,
                    "level_name": level.Name if hasattr(level, "Name") else "?",
                    "status": "skipped",
                    "reason": "source_empty",
                }
            )
            continue

        result = set_param_value(level, PARAM_TARGET_NAME, source_value)
        status = result.get("status")
        reason = result.get("reason")

        if status == "updated":
            stats["total_updated"] += 1
            stats["results"].append(
                {
                    "id": level_id,
                    "level_name": level.Name if hasattr(level, "Name") else "?",
                    "status": "updated",
                    "reason": None,
                }
            )
        elif status == "already_ok":
            stats["skip_reasons"]["already_ok"] += 1
            stats["total_skipped"] += 1
        elif reason in stats["skip_reasons"]:
            stats["skip_reasons"][reason] += 1
            stats["total_skipped"] += 1
            stats["results"].append(
                {
                    "id": level_id,
                    "level_name": level.Name if hasattr(level, "Name") else "?",
                    "status": "skipped",
                    "reason": reason,
                }
            )
        else:
            stats["skip_reasons"]["exception"] += 1
            stats["total_skipped"] += 1
            stats["results"].append(
                {
                    "id": level_id,
                    "level_name": level.Name if hasattr(level, "Name") else "?",
                    "status": "skipped",
                    "reason": reason,
                }
            )

    return stats


def print_report(stats):
    output.print_md("## Итоги копирования названий уровней")
    output.print_md(
        "> Всего уровней: **{}** | Обновлено: **{}** | Пропущено: **{}**".format(
            stats["total_processed"], stats["total_updated"], stats["total_skipped"]
        )
    )

    output.print_md("### Почему уровни пропускались")
    has_skips = False
    skip_labels = {
        "source_empty": "Пустое значение в ДГП_Имя уровня",
        "parameter_not_found": "Не найден целевой параметр",
        "readonly": "Параметр только для чтения",
        "already_ok": "Значение уже совпадает",
        "wrong_storage_type": "Неверный тип параметра",
        "exception": "Ошибка при записи",
    }
    for reason, count in stats["skip_reasons"].items():
        if count > 0:
            has_skips = True
            output.print_md(
                "- **{}**: {}".format(skip_labels.get(reason, reason), count)
            )
    if not has_skips:
        output.print_md("- Пропусков нет")

    results_table = [
        row for row in stats["results"] if row.get("status") in ["updated", "skipped"]
    ]
    if results_table:
        output.print_md("### Детали по уровням (обновленные и пропущенные)")
        output.print_md(
            '<div style="color: red; font-weight: bold;">Все ID кликабельные</div>'
        )

        def make_id_link(elem_id):
            try:
                return output.linkify(DB.ElementId(elem_id))
            except Exception:
                return to_unicode(elem_id)

        table_rows = []
        for r in results_table[:100]:
            status_label = {
                "updated": "✓ Обновлено",
                "skipped": "✗ Пропущено",
            }.get(r.get("status", ""), to_unicode(r.get("status", "?")))

            reason_label = skip_labels.get(
                r.get("reason", ""), to_unicode(r.get("reason", "-"))
            )

            table_rows.append(
                [
                    make_id_link(r.get("id", "?")),
                    r.get("level_name", "?"),
                    status_label,
                    reason_label,
                ]
            )

        output.print_table(
            table_data=table_rows,
            columns=["ID", "Уровень", "Результат", "Причина"],
            formats=["{}", "{}", "{}", "{}"],
        )

        if len(results_table) > 100:
            output.print_md(
                "- Показано первых 100 из {} результатов".format(len(results_table))
            )
    else:
        output.print_md("### Детали по уровням")
        output.print_md("- Нет данных")


def main():
    t = DB.Transaction(doc, "ДГП Уровни")
    t.Start()
    try:
        stats = process_levels()
        t.Commit()
        print_report(stats)
    except Exception as e:
        try:
            if t.GetStatus() == DB.TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        print("Ошибка: {}".format(to_unicode(e)))
        output.print_md("### Ошибка")
        output.print_md("```\n{}\n```".format(to_unicode(e)))
        import traceback

        traceback.print_exc()


if __name__ == "__main__":
    main()
