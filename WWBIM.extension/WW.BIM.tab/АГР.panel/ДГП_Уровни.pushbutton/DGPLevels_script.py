# -*- coding: utf-8 -*-
"""
ДГП Уровни — обмен значениями между параметрами ДГП_Имя уровня и Имя для всех уровней.
"""

from __future__ import print_function, division

from pyrevit import revit, DB, forms, script


doc = revit.doc
output = script.get_output()


# ============================================================================
# PARAM NAMES
# ============================================================================

PARAM_1 = "ДГП_Имя уровня"
PARAM_2 = "Имя"


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
        "total_swapped": 0,
        "skip_reasons": {
            "both_empty": 0,
            "parameter_not_found": 0,
            "readonly": 0,
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
        "both_empty": "Оба параметра пустые",
        "parameter_not_found": "Не найден параметр",
        "readonly": "Параметр только для чтения",
        "exception": "Ошибка при записи",
    }

    for level in levels:
        stats["total_processed"] += 1

        try:
            level_id = level.Id.IntegerValue
        except Exception:
            level_id = -1

        param1 = get_param(level, PARAM_1)
        param2 = get_param(level, PARAM_2)

        val1_old = get_param_text(param1)
        val2_old = get_param_text(param2)

        # Пропускаем если оба параметра пустые
        if (not val1_old or val1_old.strip() == "") and (
            not val2_old or val2_old.strip() == ""
        ):
            stats["skip_reasons"]["both_empty"] += 1
            stats["total_skipped"] += 1
            stats["results"].append(
                {
                    "id": level_id,
                    "level_name": level.Name if hasattr(level, "Name") else "?",
                    "status": "skipped",
                    "reason": "both_empty",
                    "val1_old": val1_old,
                    "val2_old": val2_old,
                    "val1_new": None,
                    "val2_new": None,
                }
            )
            continue

        # Проверяем параметры
        param1_ok = (
            param1
            and not param1.IsReadOnly
            and param1.StorageType == DB.StorageType.String
        )
        param2_ok = (
            param2
            and not param2.IsReadOnly
            and param2.StorageType == DB.StorageType.String
        )

        if not param1_ok and not param2_ok:
            stats["skip_reasons"]["parameter_not_found"] += 1
            stats["total_skipped"] += 1
            stats["results"].append(
                {
                    "id": level_id,
                    "level_name": level.Name if hasattr(level, "Name") else "?",
                    "status": "skipped",
                    "reason": "parameter_not_found",
                    "val1_old": val1_old,
                    "val2_old": val2_old,
                    "val1_new": None,
                    "val2_new": None,
                }
            )
            continue

        # Выполняем обмен значениями
        val1_new = val2_old if param1_ok else val1_old
        val2_new = val1_old if param2_ok else val2_old

        # Записываем новые значения
        if param1_ok and normalize_for_compare(val1_old) != normalize_for_compare(
            val1_new
        ):
            result = set_param_value(level, PARAM_1, val1_new)
            if result.get("status") != "updated":
                val1_new = val1_old  # Не удалось записать
        else:
            val1_new = val1_old  # Нет смысла писать то же значение

        if param2_ok and normalize_for_compare(val2_old) != normalize_for_compare(
            val2_new
        ):
            result = set_param_value(level, PARAM_2, val2_new)
            if result.get("status") != "updated":
                val2_new = val2_old  # Не удалось записать
        else:
            val2_new = val2_old  # Нет смысла писать то же значение

        # Проверяем, были ли изменения
        updated = (val1_new != val1_old) or (val2_new != val2_old)

        if updated:
            stats["total_updated"] += 1
            stats["total_swapped"] += 1
            stats["results"].append(
                {
                    "id": level_id,
                    "level_name": level.Name if hasattr(level, "Name") else "?",
                    "status": "swapped",
                    "reason": None,
                    "val1_old": val1_old,
                    "val2_old": val2_old,
                    "val1_new": val1_new,
                    "val2_new": val2_new,
                }
            )
        else:
            stats["total_skipped"] += 1
            stats["results"].append(
                {
                    "id": level_id,
                    "level_name": level.Name if hasattr(level, "Name") else "?",
                    "status": "skipped",
                    "reason": "no_change",
                    "val1_old": val1_old,
                    "val2_old": val2_old,
                    "val1_new": val1_new,
                    "val2_new": val2_new,
                }
            )

    return stats


def print_report(stats):
    output.print_md("## Итоги обмена значениями уровней")
    output.print_md(
        "> Всего уровней: **{}** | Обменено: **{}** | Пропущено: **{}**".format(
            stats["total_processed"], stats["total_updated"], stats["total_skipped"]
        )
    )

    output.print_md("### Почему уровни пропускались")
    has_skips = False
    skip_labels = {
        "both_empty": "Оба параметра пустые",
        "parameter_not_found": "Не найден параметр",
        "readonly": "Параметр только для чтения",
        "exception": "Ошибка при записи",
        "no_change": "Значения уже одинаковые",
    }
    for reason, count in stats["skip_reasons"].items():
        if count > 0:
            has_skips = True
            output.print_md(
                "- **{}**: {}".format(skip_labels.get(reason, reason), count)
            )
    if stats["skip_reasons"].get("no_change", 0) > 0:
        has_skips = True
        output.print_md(
            "- **{}**: {}".format(
                skip_labels["no_change"], stats["skip_reasons"]["no_change"]
            )
        )
    if not has_skips:
        output.print_md("- Пропусков нет")

    results_table = [
        row for row in stats["results"] if row.get("status") in ["swapped", "skipped"]
    ]
    if results_table:
        output.print_md("### Детали по уровням (обмененные и пропущенные)")
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
                "swapped": "✓ Обменено",
                "skipped": "✗ Пропущено",
            }.get(r.get("status", ""), to_unicode(r.get("status", "?")))

            reason_label = skip_labels.get(r.get("reason", ""), "-")
            if r.get("reason") == "no_change":
                reason_label = skip_labels["no_change"]

            val1_old = r.get("val1_old", "")
            val2_old = r.get("val2_old", "")
            val1_new = r.get("val1_new", "")
            val2_new = r.get("val2_new", "")

            # Формируем красивый вывод обмена
            if r.get("status") == "swapped":
                change1 = "{} → {}".format(val1_old, val1_new)
                change2 = "{} → {}".format(val2_old, val2_new)
            else:
                change1 = val1_old
                change2 = val2_old

            table_rows.append(
                [
                    make_id_link(r.get("id", "?")),
                    r.get("level_name", "?"),
                    status_label,
                    reason_label,
                    change1,
                    change2,
                ]
            )

        output.print_table(
            table_data=table_rows,
            columns=[
                "ID",
                "Уровень",
                "Результат",
                "Причина",
                "{} → {}".format(PARAM_1),
                "{} → {}".format(PARAM_2),
            ],
            formats=["{}", "{}", "{}", "{}", "{}", "{}"],
        )

        if len(results_table) > 100:
            output.print_md(
                "- Показано первых 100 из {} результатов".format(len(results_table))
            )
    else:
        output.print_md("### Детали по уровням")
        output.print_md("- Нет данных")


def main():
    t = DB.Transaction(doc, "ДГП Уровни: обмен значениями")
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
