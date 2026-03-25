# -*- coding: utf-8 -*-
"""
ДГП Параметры - заполнение параметров экземпляра по категории и условию.

Правила:
- Для большинства категорий условие берется из параметра "Группа модели"
- Для категории Зона (OST_Areas) условие берется из имени схемы зонирования
"""
from __future__ import print_function, division

import re
from pyrevit import revit, DB, forms, script


doc = revit.doc
output = script.get_output()


# ============================================================================
# PARAM NAMES
# ============================================================================

PARAM_GROUP_MODEL = u"Группа модели"

PARAM_TARGET_LAND_AREA = u"ДГП_Площадь земельного участка"
PARAM_TARGET_AREA = u"ДГП_Площадь"
PARAM_TARGET_LENGTH = u"ДГП_Длина"
PARAM_TARGET_THICKNESS = u"ДГП_Толщина"
PARAM_TARGET_HEIGHT = u"ДГП_Высота"
PARAM_TARGET_WIDTH = u"ДГП_Ширина"
PARAM_TARGET_DIAMETER = u"ДГП_Диаметр"
PARAM_TARGET_CODE = u"ДГП_Код помещения и зоны МССК"
PARAM_TARGET_NAME = u"ДГП_Наименование помещения и зоны МССК"
PARAM_TARGET_PROJECT_NAME = u"ДГП_Наименование проекта"
PARAM_TARGET_PROJECT_ADDRESS = u"ДГП_Адрес"

PARAM_SOURCE_AREA = u"Площадь"
PARAM_SOURCE_AREA_COEF = u"ADSK_Площадь с коэффициентом"
PARAM_SOURCE_HEIGHT = u"ADSK_Размер_Высота"
PARAM_SOURCE_WIDTH = u"ADSK_Размер_Ширина"
PARAM_SOURCE_DIAMETER = u"ADSK_Размер_Диаметр"
PARAM_SOURCE_THICKNESS = u"Толщина"
PARAM_SOURCE_PROJECT_NAME = u"Наименование проекта"
PARAM_SOURCE_PROJECT_ADDRESS = u"Адрес проекта"


# ============================================================================
# HELPERS
# ============================================================================

UNIT_SUFFIX_RE = re.compile(
    ur"\s*[\u00A0\u0020]*(?:м²|м2|кв\.?м|кв\.?м²|квм|мм|см|м|м\.?п\.?|п\.?м\.?)\s*$",
    re.IGNORECASE,
)


def to_unicode(v):
    if v is None:
        return u""
    try:
        return unicode(v)
    except Exception:
        try:
            return str(v)
        except Exception:
            return u""


def bic_to_int(bic):
    """Безопасно получить IntegerValue из BuiltInCategory enum."""
    try:
        return int(bic)
    except Exception:
        try:
            cat = DB.Category.GetCategory(doc, bic)
            if cat:
                return cat.Id.IntegerValue
        except Exception:
            pass
    return None


def normalize_for_compare(v):
    s = to_unicode(v).replace(u"\u00A0", u" ").strip()
    return re.sub(ur"\s+", u" ", s)


def strip_unit_suffix(v):
    s = normalize_for_compare(v)
    if not s:
        return u""
    return UNIT_SUFFIX_RE.sub(u"", s).strip()


def parse_float(v):
    if v is None:
        return None

    if isinstance(v, float):
        return v
    if isinstance(v, int):
        return float(v)

    s = strip_unit_suffix(v)
    if not s:
        return None

    s = s.replace(u" ", u"").replace(u",", u".")
    try:
        return float(s)
    except Exception:
        return None


def get_param(elem, name):
    if not elem or not name:
        return None
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None


def get_type_param(elem, name):
    try:
        type_id = elem.GetTypeId()
        if not type_id or type_id == DB.ElementId.InvalidElementId:
            return None
        elem_type = elem.Document.GetElement(type_id)
        if not elem_type:
            return None
        return elem_type.LookupParameter(name)
    except Exception:
        return None


def has_adsk_dimension_param(elem, param_name):
    """Проверяет наличие ADSK_Размер_ параметра в экземпляре или типе."""
    p_inst = get_param(elem, param_name)
    if p_inst:
        return True

    p_type = get_type_param(elem, param_name)
    return p_type is not None


def get_param_text(param):
    if not param or not param.HasValue:
        return None
    try:
        vs = param.AsValueString()
        if vs:
            return strip_unit_suffix(vs)
    except Exception:
        pass
    try:
        if param.StorageType == DB.StorageType.String:
            return strip_unit_suffix(param.AsString())
        if param.StorageType == DB.StorageType.Double:
            return strip_unit_suffix(param.AsValueString())
        if param.StorageType == DB.StorageType.Integer:
            return to_unicode(param.AsInteger())
        if param.StorageType == DB.StorageType.ElementId:
            eid = param.AsElementId()
            if eid:
                return to_unicode(eid.IntegerValue)
    except Exception:
        return None
    return None


def get_param_number(param):
    if not param or not param.HasValue:
        return None
    try:
        st = param.StorageType
        if st == DB.StorageType.Double:
            return param.AsDouble()
        if st == DB.StorageType.Integer:
            return float(param.AsInteger())
        if st == DB.StorageType.String:
            return parse_float(param.AsString())
        vs = param.AsValueString()
        return parse_float(vs)
    except Exception:
        return None


def resolve_text(elem, source_spec):
    if not source_spec:
        return None
    if source_spec.startswith(u"const:"):
        return source_spec[6:]

    use_type = source_spec.startswith(u"type:")
    name = source_spec[5:] if use_type else source_spec

    p = get_type_param(elem, name) if use_type else get_param(elem, name)
    if p is None and not use_type:
        p = get_type_param(elem, name)
    return get_param_text(p)


def resolve_number(elem, source_spec):
    if not source_spec:
        return None
    if source_spec.startswith(u"const:"):
        return parse_float(source_spec[6:])

    use_type = source_spec.startswith(u"type:")
    name = source_spec[5:] if use_type else source_spec

    p = get_type_param(elem, name) if use_type else get_param(elem, name)
    if p is None and not use_type:
        p = get_type_param(elem, name)
    return get_param_number(p)


def should_write_text(current, expected):
    return normalize_for_compare(current) != normalize_for_compare(expected)


def should_write_number(current, expected):
    if current is None:
        return True
    if expected is None:
        return False
    try:
        return abs(float(current) - float(expected)) > 1e-9
    except Exception:
        return True


def get_group_value_for_rule(elem, cat_bic):
    # Areas: условие по имени схемы зонирования
    if cat_bic == bic_to_int(DB.BuiltInCategory.OST_Areas):
        try:
            scheme = elem.AreaScheme
            if scheme and scheme.Name:
                return normalize_for_compare(scheme.Name)
        except Exception:
            pass
        return u""

    # Остальные категории: условие по "Группа модели" (только в типе)
    p = get_type_param(elem, PARAM_GROUP_MODEL)
    return normalize_for_compare(get_param_text(p))


def find_rule_for_element(elem, rules):
    try:
        cat = elem.Category
        if not cat:
            return None
        cat_bic = cat.Id.IntegerValue
    except Exception:
        return None

    group_val = get_group_value_for_rule(elem, cat_bic)
    default_rule = None

    for rule in rules:
        rule_cat = bic_to_int(rule.get(u"category"))
        if rule_cat != cat_bic:
            continue

        rg = rule.get(u"group")
        rgs = rule.get(u"groups")

        if rg is not None:
            if normalize_for_compare(rg) == group_val:
                return rule
            continue

        if rgs is not None:
            for v in rgs:
                if normalize_for_compare(v) == group_val:
                    return rule
            continue

        if default_rule is None:
            default_rule = rule

    return default_rule


def set_param_from_assignment(elem, assignment, missing_adsk_callback=None):
    target_name = assignment.get(u"target")
    source_spec = assignment.get(u"source")
    unit_type = assignment.get(u"unit_type")  # "number" or "text"

    if not target_name or not source_spec:
        return {
            u"status": u"skipped",
            u"reason": u"invalid_assignment",
            u"target": target_name,
            u"source": source_spec,
        }

    # Проверка наличия ADSK_Размер_ параметров
    if source_spec.startswith(u"ADSK_Размер_"):
        param_to_check = source_spec[5:] if source_spec.startswith(u"type:") else source_spec
        if not has_adsk_dimension_param(elem, param_to_check):
            result = {
                u"status": u"skipped",
                u"reason": u"missing_adsk_param",
                u"target": target_name,
                u"source": source_spec,
            }
            if missing_adsk_callback:
                missing_adsk_callback(elem, param_to_check)
            return result

    target = get_param(elem, target_name)
    if target is None:
        return {
            u"status": u"parameter_not_found",
            u"reason": u"parameter_not_found",
            u"target": target_name,
            u"source": source_spec,
        }
    if target.IsReadOnly:
        return {
            u"status": u"readonly",
            u"reason": u"readonly",
            u"target": target_name,
            u"source": source_spec,
        }

    st = target.StorageType

    # Текстовые параметры
    if st == DB.StorageType.String:
        expected = resolve_text(elem, source_spec)
        if expected is None or expected == u"":
            return {
                u"status": u"skipped",
                u"reason": u"source_empty",
                u"target": target_name,
                u"source": source_spec,
            }
        current = target.AsString()
        if not should_write_text(current, expected):
            return {
                u"status": u"already_ok",
                u"reason": u"already_ok",
                u"target": target_name,
                u"source": source_spec,
            }
        try:
            target.Set(to_unicode(expected))
            return {
                u"status": u"updated",
                u"reason": None,
                u"target": target_name,
                u"source": source_spec,
            }
        except Exception as e:
            return {
                u"status": u"exception",
                u"reason": to_unicode(e),
                u"target": target_name,
                u"source": source_spec,
            }

    # Числовые параметры
    if st == DB.StorageType.Double:
        expected = resolve_number(elem, source_spec)
        if expected is None:
            return {
                u"status": u"skipped",
                u"reason": u"source_empty",
                u"target": target_name,
                u"source": source_spec,
            }
        current = target.AsDouble()
        if not should_write_number(current, expected):
            return {
                u"status": u"already_ok",
                u"reason": u"already_ok",
                u"target": target_name,
                u"source": source_spec,
            }
        try:
            target.Set(float(expected))
            return {
                u"status": u"updated",
                u"reason": None,
                u"target": target_name,
                u"source": source_spec,
            }
        except Exception as e:
            return {
                u"status": u"exception",
                u"reason": to_unicode(e),
                u"target": target_name,
                u"source": source_spec,
            }

    return {
        u"status": u"wrong_storage_type",
        u"reason": u"wrong_storage_type",
        u"target": target_name,
        u"source": source_spec,
    }


# ============================================================================
# RULES
# ============================================================================

RULES = [
    # Топография
    {
        u"category": DB.BuiltInCategory.OST_Topography,
        u"assignments": [
            {u"target": PARAM_TARGET_LAND_AREA, u"source": PARAM_SOURCE_AREA},
        ],
    },

    # Перекрытия - Зонирование застройки
    {
        u"category": DB.BuiltInCategory.OST_Floors,
        u"group": u"Зонирование застройки",
        u"assignments": [
            {u"target": PARAM_TARGET_AREA, u"source": PARAM_SOURCE_AREA},
        ],
    },

    # Перекрытия - default
    {
        u"category": DB.BuiltInCategory.OST_Floors,
        u"assignments": [
            {u"target": PARAM_TARGET_THICKNESS, u"source": u"type:" + PARAM_SOURCE_THICKNESS},
        ],
    },

    # Окна - Заполнение оконных проёмов
    {
        u"category": DB.BuiltInCategory.OST_Windows,
        u"group": u"Заполнение оконных проёмов",
        u"assignments": [
            {u"target": PARAM_TARGET_HEIGHT, u"source": PARAM_SOURCE_HEIGHT},
            {u"target": PARAM_TARGET_WIDTH, u"source": PARAM_SOURCE_WIDTH},
        ],
    },

    # Стена - Заполнение оконных проёмов
    {
        u"category": DB.BuiltInCategory.OST_Walls,
        u"group": u"Заполнение оконных проёмов",
        u"assignments": [
            {u"target": PARAM_TARGET_HEIGHT, u"source": PARAM_SOURCE_HEIGHT},
            {u"target": PARAM_TARGET_WIDTH, u"source": PARAM_SOURCE_WIDTH},
        ],
    },

    # Стены - Фасад
    {
        u"category": DB.BuiltInCategory.OST_Walls,
        u"group": u"Фасад",
        u"assignments": [
            {u"target": PARAM_TARGET_THICKNESS, u"source": u"type:" + PARAM_SOURCE_THICKNESS},
        ],
    },

    # Стены - default
    {
        u"category": DB.BuiltInCategory.OST_Walls,
        u"assignments": [
            {u"target": PARAM_TARGET_THICKNESS, u"source": u"type:" + PARAM_SOURCE_THICKNESS},
        ],
    },

    # Двери
    {
        u"category": DB.BuiltInCategory.OST_Doors,
        u"groups": [u"Дверь", u"Ворота", u"Люк"],
        u"assignments": [
            {u"target": PARAM_TARGET_HEIGHT, u"source": PARAM_SOURCE_HEIGHT},
            {u"target": PARAM_TARGET_WIDTH, u"source": PARAM_SOURCE_WIDTH},
        ],
    },

    # Несущие колонны
    {
        u"category": DB.BuiltInCategory.OST_StructuralColumns,
        u"assignments": [
            {u"target": PARAM_TARGET_HEIGHT, u"source": PARAM_SOURCE_HEIGHT},
            {u"target": PARAM_TARGET_WIDTH, u"source": PARAM_SOURCE_WIDTH},
            {u"target": PARAM_TARGET_DIAMETER, u"source": PARAM_SOURCE_DIAMETER},
        ],
    },

    # Помещения
    {
        u"category": DB.BuiltInCategory.OST_Rooms,
        u"assignments": [
            {u"target": PARAM_TARGET_AREA, u"source": PARAM_SOURCE_AREA_COEF},
        ],
    },

    # Зоны (условие по имени схемы)
    {
        u"category": DB.BuiltInCategory.OST_Areas,
        u"group": u"СПП в ГНС",
        u"assignments": [
            {u"target": PARAM_TARGET_CODE, u"source": u"const:9999"},
            {u"target": PARAM_TARGET_NAME, u"source": u"const:9999"},
            {u"target": PARAM_TARGET_AREA, u"source": PARAM_SOURCE_AREA},
        ],
    },
    {
        u"category": DB.BuiltInCategory.OST_Areas,
        u"group": u"Общая площадь",
        u"assignments": [
            {u"target": PARAM_TARGET_CODE, u"source": u"const:П3 03"},
            {u"target": PARAM_TARGET_NAME, u"source": u"const:9999"},
            {u"target": PARAM_TARGET_AREA, u"source": PARAM_SOURCE_AREA},
        ],
    },

    # Сведения о проекте
    {
        u"category": DB.BuiltInCategory.OST_ProjectInformation,
        u"assignments": [
            {u"target": PARAM_TARGET_PROJECT_NAME, u"source": PARAM_SOURCE_PROJECT_NAME},
            {u"target": PARAM_TARGET_PROJECT_ADDRESS, u"source": PARAM_SOURCE_PROJECT_ADDRESS},
        ],
    },
]


def process_elements():
    stats = {
        u"total_processed": 0,
        u"total_updated": 0,
        u"total_skipped": 0,
        u"by_category": {},
        u"skip_reasons": {
            u"parameter_not_found": 0,
            u"wrong_storage_type": 0,
            u"readonly": 0,
            u"already_ok": 0,
            u"source_empty": 0,
            u"exception": 0,
            u"no_rule": 0,
            u"missing_adsk_param": 0,
        },
        u"errors": [],
        u"by_target": {},
        u"debug_samples": [],
        u"no_rule_samples": [],
        u"missing_adsk_samples": [],
    }

    def touch_target(target_name):
        if target_name not in stats[u"by_target"]:
            stats[u"by_target"][target_name] = {
                u"updated": 0,
                u"already_ok": 0,
                u"source_empty": 0,
                u"parameter_not_found": 0,
                u"wrong_storage_type": 0,
                u"readonly": 0,
                u"exception": 0,
                u"invalid_assignment": 0,
            }

    def add_sample(elem, cat_name, result):
        if len(stats[u"debug_samples"]) >= 60:
            return
        try:
            elem_id = elem.Id.IntegerValue
        except Exception:
            elem_id = -1
        stats[u"debug_samples"].append(
            {
                u"id": elem_id,
                u"category": cat_name,
                u"target": result.get(u"target"),
                u"source": result.get(u"source"),
                u"status": result.get(u"status"),
                u"reason": result.get(u"reason"),
            }
        )

    def add_missing_adsk_sample(elem, missing_param):
        """Добавляет элемент без ADSK_Размер_ параметра в статистику."""
        if len(stats[u"missing_adsk_samples"]) >= 200:
            return
        try:
            elem_id = elem.Id.IntegerValue
        except Exception:
            elem_id = -1
        try:
            cat_name = to_unicode(elem.Category.Name)
        except Exception:
            cat_name = u"?"
        stats[u"missing_adsk_samples"].append(
            {
                u"elem": elem,
                u"id": elem_id,
                u"category": cat_name,
                u"missing_param": missing_param,
            }
        )

    def get_category_adsk_params(cat_bic):
        """Возвращает ADSK_Размер_ параметры, используемые правилами категории."""
        params = []
        for r in RULES:
            if r.get(u"category") != cat_bic:
                continue
            for a in r.get(u"assignments", []):
                src = a.get(u"source")
                if not src:
                    continue
                if src.startswith(u"type:"):
                    src_name = src[5:]
                else:
                    src_name = src
                if src_name.startswith(u"ADSK_Размер_") and src_name not in params:
                    params.append(src_name)
        return params

    def expected_rules_for_category(cat_bic):
        vals = []
        has_default = False
        for r in RULES:
            if r.get(u"category") != cat_bic:
                continue
            if r.get(u"group") is not None:
                vals.append(to_unicode(r.get(u"group")))
            elif r.get(u"groups") is not None:
                groups = r.get(u"groups") or []
                for g in groups:
                    vals.append(to_unicode(g))
            else:
                has_default = True
        if has_default:
            vals.append(u"<default>")
        return u", ".join(vals) if vals else u"<none>"

    categories = set([r.get(u"category") for r in RULES if r.get(u"category")])

    for category in categories:
        try:
            cat_name = to_unicode(category).replace(u"OST_", u"")
        except Exception:
            cat_name = to_unicode(category)

        stats[u"by_category"][cat_name] = {u"processed": 0, u"updated": 0, u"skipped": 0}

        try:
            elements = (
                DB.FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception as e:
            stats[u"errors"].append({u"category": cat_name, u"error": to_unicode(e)})
            continue

        for elem in elements:
            stats[u"total_processed"] += 1
            stats[u"by_category"][cat_name][u"processed"] += 1

            rule = find_rule_for_element(elem, RULES)
            if not rule:
                # Даже если правило не найдено, проверяем ADSK_Размер_ параметры категории
                for adsk_name in get_category_adsk_params(category):
                    if not has_adsk_dimension_param(elem, adsk_name):
                        add_missing_adsk_sample(elem, adsk_name)

                stats[u"skip_reasons"][u"no_rule"] += 1
                stats[u"total_skipped"] += 1
                stats[u"by_category"][cat_name][u"skipped"] += 1
                if len(stats[u"no_rule_samples"]) < 60:
                    try:
                        elem_id = elem.Id.IntegerValue
                    except Exception:
                        elem_id = -1
                    try:
                        actual_cat_id = elem.Category.Id.IntegerValue
                    except Exception:
                        actual_cat_id = None
                    actual_group = get_group_value_for_rule(elem, actual_cat_id)
                    stats[u"no_rule_samples"].append(
                        {
                            u"id": elem_id,
                            u"category": cat_name,
                            u"actual_group": to_unicode(actual_group),
                            u"expected": expected_rules_for_category(category),
                        }
                    )
                continue

            elem_updated = False
            for assignment in rule.get(u"assignments", []):
                result = set_param_from_assignment(elem, assignment, add_missing_adsk_sample)
                status = result.get(u"status")
                reason = result.get(u"reason")
                target_name = result.get(u"target")

                if target_name:
                    touch_target(target_name)

                if status == u"updated":
                    elem_updated = True
                    if target_name:
                        stats[u"by_target"][target_name][u"updated"] += 1
                elif status == u"already_ok":
                    stats[u"skip_reasons"][u"already_ok"] += 1
                    if target_name:
                        stats[u"by_target"][target_name][u"already_ok"] += 1
                elif reason in stats[u"skip_reasons"]:
                    stats[u"skip_reasons"][reason] += 1
                    if target_name and reason in stats[u"by_target"][target_name]:
                        stats[u"by_target"][target_name][reason] += 1
                    add_sample(elem, cat_name, result)
                else:
                    add_sample(elem, cat_name, result)

            if elem_updated:
                stats[u"total_updated"] += 1
                stats[u"by_category"][cat_name][u"updated"] += 1
            else:
                stats[u"total_skipped"] += 1
                stats[u"by_category"][cat_name][u"skipped"] += 1

    return stats


def print_report(stats):
    skip_labels = {
        u"parameter_not_found": u"Не найден целевой параметр",
        u"wrong_storage_type": u"Неверный тип целевого параметра",
        u"readonly": u"Целевой параметр только для чтения",
        u"already_ok": u"Значение уже было заполнено",
        u"source_empty": u"Пустое исходное значение",
        u"exception": u"Ошибка при записи значения",
        u"no_rule": u"Не найдено подходящее правило",
        u"missing_adsk_param": u"Отсутствует параметр ADSK_Размер_",
    }
    status_labels = {
        u"updated": u"Обновлено",
        u"already_ok": u"Без изменений",
        u"skipped": u"Пропущено",
        u"parameter_not_found": u"Параметр не найден",
        u"readonly": u"Только чтение",
        u"wrong_storage_type": u"Неверный тип",
        u"exception": u"Ошибка",
    }

    def make_id_link(elem_id):
        try:
            if isinstance(elem_id, int):
                return output.linkify(DB.ElementId(elem_id))
            return output.linkify(DB.ElementId(int(elem_id)))
        except Exception:
            return to_unicode(elem_id)

    output.print_md(u"## Итоги заполнения ДГП")
    output.print_md(
        u"> Всего элементов: **{}** | Успешно заполнено: **{}** | Пропущено: **{}**".format(
            stats[u"total_processed"], stats[u"total_updated"], stats[u"total_skipped"]
        )
    )

    output.print_md(u"### Что получилось по категориям")
    sorted_categories = sorted(
        stats[u"by_category"].items(),
        key=lambda kv: kv[1].get(u"processed", 0),
        reverse=True,
    )
    category_rows = []
    for cat_name, cat_stats in sorted_categories:
        category_rows.append(
            [
                cat_name,
                cat_stats.get(u"processed", 0),
                cat_stats.get(u"updated", 0),
                cat_stats.get(u"skipped", 0),
            ]
        )
    output.print_table(
        table_data=category_rows,
        columns=[u"Категория", u"Обработано", u"Заполнено", u"Пропущено"],
        formats=[u"{}", u"{}", u"{}", u"{}"],
    )

    output.print_md(u"### Почему элементы пропускались")
    has_skips = False
    for reason, count in stats[u"skip_reasons"].items():
        if count > 0:
            has_skips = True
            output.print_md(u"- **{}**: {}".format(skip_labels.get(reason, reason), count))
    if not has_skips:
        output.print_md(u"- Пропусков нет")

    output.print_md(u"### Как заполнялись целевые параметры")
    if not stats[u"by_target"]:
        output.print_md(u"- Данных нет")
    else:
        target_rows = []
        for target_name, t in sorted(stats[u"by_target"].items()):
            target_rows.append(
                [
                    target_name,
                    t.get(u"updated", 0),
                    t.get(u"already_ok", 0),
                    t.get(u"source_empty", 0),
                    t.get(u"parameter_not_found", 0),
                ]
            )
        output.print_table(
            table_data=target_rows,
            columns=[u"Параметр", u"Записано", u"Без изменений", u"Пустой источник", u"Не найден"],
            formats=[u"{}", u"{}", u"{}", u"{}", u"{}"],
        )

    # Единый раздел проблем: объединяем no_rule + debug + missing ADSK
    problem_rows = []

    for s in stats.get(u"no_rule_samples", []):
        problem_rows.append([
            make_id_link(s.get(u"id", u"?")),
            s.get(u"category", u"?"),
            u"Не найдено правило",
            u"Текущая группа: '{}'. Ожидалось: '{}'".format(
                s.get(u"actual_group", u""),
                s.get(u"expected", u""),
            ),
        ])

    for s in stats.get(u"debug_samples", []):
        problem_rows.append([
            make_id_link(s.get(u"id", u"?")),
            s.get(u"category", u"?"),
            status_labels.get(s.get(u"status", u""), to_unicode(s.get(u"status", u"?"))),
            u"Куда: '{}' | Откуда: '{}' | Причина: '{}'".format(
                s.get(u"target", u"?"),
                s.get(u"source", u"?"),
                skip_labels.get(s.get(u"reason", u""), to_unicode(s.get(u"reason", u"-"))),
            ),
        ])

    for s in stats.get(u"missing_adsk_samples", []):
        problem_rows.append([
            make_id_link(s.get(u"id", u"?")),
            s.get(u"category", u"?"),
            u"Нет ADSK_Размер_",
            u"Не найден параметр: '{}'".format(s.get(u"missing_param", u"?")),
        ])

    if problem_rows:
        output.print_md(u'### Проблемные элементы')
        output.print_md(u'<div style="color: red; font-weight: bold;">Все ID в таблице кликабельные</div>')
        output.print_table(
            table_data=problem_rows,
            columns=[u"ID элемента", u"Категория", u"Ситуация", u"Подробности"],
            formats=[u"{}", u"{}", u"{}", u"{}"],
        )
    else:
        output.print_md(u"### Проблемные элементы")
        output.print_md(u"- Проблем не обнаружено")

    if stats[u"errors"]:
        output.print_md(u"### Ошибки выполнения")
        for err in stats[u"errors"]:
            output.print_md(u"- Категория **{}**: {}".format(err.get(u"category", u"?"), err.get(u"error", u"?")))


def main():
    t = DB.Transaction(doc, u"ДГП Параметры")
    t.Start()
    try:
        stats = process_elements()
        t.Commit()
        print_report(stats)
    except Exception as e:
        try:
            if t.GetStatus() == DB.TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        print("Ошибка: {}".format(to_unicode(e)))
        output.print_md(u"### Ошибка")
        output.print_md(u"```\n{}\n```".format(to_unicode(e)))
        import traceback

        traceback.print_exc()


if __name__ == u"__main__":
    main()
