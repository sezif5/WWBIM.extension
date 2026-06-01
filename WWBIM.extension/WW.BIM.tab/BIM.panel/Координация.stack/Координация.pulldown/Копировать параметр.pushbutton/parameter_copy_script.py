# -*- coding: utf-8 -*-
from __future__ import print_function, division

import re

from pyrevit import revit, DB, forms, script


doc = revit.doc
output = script.get_output()


SCOPE_INSTANCE = u"Экземпляр"
SCOPE_TYPE = u"Тип"


class ParamChoice(object):
    def __init__(self, name, scope):
        self.name = name
        self.scope = scope
        self.label = u"{} | {}".format(scope, name)


def to_unicode(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def normalize_text(value):
    return re.sub(u"\\s+", u" ", to_unicode(value).replace(u"\u00A0", u" ")).strip()


def get_selected_elements():
    try:
        return [el for el in revit.get_selection() if el is not None]
    except Exception:
        return []


def get_element_type(elem):
    try:
        type_id = elem.GetTypeId()
        if not type_id or type_id == DB.ElementId.InvalidElementId:
            return None
        return doc.GetElement(type_id)
    except Exception:
        return None


def get_param_owner(elem, scope):
    if scope == SCOPE_TYPE:
        return get_element_type(elem)
    return elem


def get_param_by_choice(elem, choice):
    try:
        owner = get_param_owner(elem, choice.scope)
        if owner is None:
            return None, None
        return owner.LookupParameter(choice.name), owner
    except Exception:
        return None, None


def get_category_name(elem):
    try:
        if elem.Category:
            return to_unicode(elem.Category.Name)
    except Exception:
        pass
    return u"?"


def get_element_id(elem):
    try:
        return elem.Id.IntegerValue
    except Exception:
        return -1


def get_owner_id(owner):
    try:
        return owner.Id.IntegerValue
    except Exception:
        return -1


def get_param_storage_label(param):
    try:
        st = param.StorageType
        if st == DB.StorageType.String:
            return u"Текст"
        if st == DB.StorageType.Double:
            return u"Число"
        if st == DB.StorageType.Integer:
            return u"Целое"
        if st == DB.StorageType.ElementId:
            return u"ElementId"
    except Exception:
        pass
    return u"?"


def get_param_spec_key(param):
    try:
        definition = param.Definition
    except Exception:
        return None
    try:
        if hasattr(definition, "GetDataType"):
            return to_unicode(definition.GetDataType().TypeId)
    except Exception:
        pass
    try:
        return to_unicode(definition.ParameterType)
    except Exception:
        return None


def spec_is(param, spec_name):
    key = get_param_spec_key(param) or u""
    key_low = key.lower()
    if spec_name == u"length":
        return u"length" in key_low or key in (u"Length", u"Длина")
    if spec_name == u"area":
        return u"area" in key_low or key in (u"Area", u"Площадь")
    if spec_name == u"volume":
        return u"volume" in key_low or key in (u"Volume", u"Объем", u"Объём")
    if spec_name == u"number":
        return u"number" in key_low or key in (u"Number", u"Число")
    return False


def collect_parameters(elements):
    choices = {}
    for elem in elements:
        try:
            for param in elem.Parameters:
                try:
                    name = to_unicode(param.Definition.Name)
                    if name:
                        choices[(SCOPE_INSTANCE, name)] = ParamChoice(name, SCOPE_INSTANCE)
                except Exception:
                    pass
        except Exception:
            pass

        elem_type = get_element_type(elem)
        if elem_type is None:
            continue
        try:
            for param in elem_type.Parameters:
                try:
                    name = to_unicode(param.Definition.Name)
                    if name:
                        choices[(SCOPE_TYPE, name)] = ParamChoice(name, SCOPE_TYPE)
                except Exception:
                    pass
        except Exception:
            pass

    return sorted(choices.values(), key=lambda c: (c.scope, c.name.lower()))


def get_element_name_by_id(elem_id):
    try:
        if not elem_id or elem_id == DB.ElementId.InvalidElementId:
            return u""
        linked_elem = doc.GetElement(elem_id)
        if linked_elem:
            return to_unicode(linked_elem.Name)
    except Exception:
        pass
    try:
        return to_unicode(elem_id.IntegerValue)
    except Exception:
        return u""


def read_param_value(param):
    if param is None:
        return {u"status": u"missing"}
    try:
        if not param.HasValue:
            return {u"status": u"empty"}
    except Exception:
        return {u"status": u"empty"}

    try:
        st = param.StorageType
        if st == DB.StorageType.String:
            return {u"status": u"ok", u"storage": st, u"raw": param.AsString(), u"text": param.AsString()}
        if st == DB.StorageType.Double:
            value_text = None
            try:
                value_text = param.AsValueString()
            except Exception:
                value_text = None
            return {
                u"status": u"ok",
                u"storage": st,
                u"raw": param.AsDouble(),
                u"text": value_text,
                u"spec_key": get_param_spec_key(param),
            }
        if st == DB.StorageType.Integer:
            value_int = param.AsInteger()
            value_text = None
            try:
                value_text = param.AsValueString()
            except Exception:
                value_text = None
            return {u"status": u"ok", u"storage": st, u"raw": value_int, u"text": value_text or to_unicode(value_int)}
        if st == DB.StorageType.ElementId:
            elem_id = param.AsElementId()
            return {u"status": u"ok", u"storage": st, u"raw": elem_id, u"text": get_element_name_by_id(elem_id)}
    except Exception as ex:
        return {u"status": u"exception", u"reason": to_unicode(ex)}

    return {u"status": u"unsupported"}


def parse_number_from_text(value):
    text = normalize_text(value)
    if not text:
        return None
    match = re.search(u"[-+]?\\d+(?:[\\s\\u00A0]*\\d{3})*(?:[\\.,]\\d+)?|[-+]?\\d+(?:[\\.,]\\d+)?", text)
    if not match:
        return None
    number_text = match.group(0).replace(u" ", u"").replace(u"\u00A0", u"").replace(u",", u".")
    try:
        return float(number_text)
    except Exception:
        return None


def get_metric_unit_for_target(target_param, value_text):
    if not hasattr(DB, "UnitTypeId"):
        return None
    text = normalize_text(value_text).lower()
    try:
        if spec_is(target_param, u"length"):
            if u"мм" in text:
                return DB.UnitTypeId.Millimeters
            if u"см" in text:
                return DB.UnitTypeId.Centimeters
            return DB.UnitTypeId.Meters
        if spec_is(target_param, u"area"):
            return DB.UnitTypeId.SquareMeters
        if spec_is(target_param, u"volume"):
            return DB.UnitTypeId.CubicMeters
    except Exception:
        return None
    return None


def convert_number_for_double_target(target_param, value_text, number_value):
    unit_id = get_metric_unit_for_target(target_param, value_text)
    if unit_id is not None:
        try:
            return DB.UnitUtils.ConvertToInternalUnits(float(number_value), unit_id)
        except Exception:
            return None
    if spec_is(target_param, u"number"):
        return float(number_value)
    return None


def values_equal(target_param, source_value):
    try:
        st = target_param.StorageType
        if st == DB.StorageType.String:
            return normalize_text(target_param.AsString()) == normalize_text(source_value.get(u"text"))
        if st == DB.StorageType.Double and source_value.get(u"storage") == DB.StorageType.Double:
            if source_value.get(u"spec_key") == get_param_spec_key(target_param):
                return abs(target_param.AsDouble() - float(source_value.get(u"raw"))) < 1e-9
            return False
        if st == DB.StorageType.Integer:
            raw = source_value.get(u"raw")
            try:
                return target_param.AsInteger() == int(raw)
            except Exception:
                number_value = parse_number_from_text(source_value.get(u"text"))
                return number_value is not None and target_param.AsInteger() == int(round(number_value))
        if st == DB.StorageType.ElementId and source_value.get(u"storage") == DB.StorageType.ElementId:
            return target_param.AsElementId() == source_value.get(u"raw")
    except Exception:
        pass
    return False


def set_string_value(target_param, source_value):
    value_text = source_value.get(u"text")
    if value_text is None or normalize_text(value_text) == u"":
        raw = source_value.get(u"raw")
        value_text = to_unicode(raw)
    target_param.Set(to_unicode(value_text))
    return True


def set_double_value(target_param, source_value):
    source_storage = source_value.get(u"storage")
    if source_storage == DB.StorageType.Double:
        if source_value.get(u"spec_key") == get_param_spec_key(target_param):
            target_param.Set(float(source_value.get(u"raw")))
            return True
        value_text = source_value.get(u"text")
        if value_text is not None and normalize_text(value_text):
            try:
                if target_param.SetValueString(to_unicode(value_text)):
                    return True
            except Exception:
                pass
            number_value = parse_number_from_text(value_text)
            if number_value is not None:
                converted_value = convert_number_for_double_target(target_param, value_text, number_value)
                if converted_value is not None:
                    target_param.Set(float(converted_value))
                    return True
        return False

    value_text = source_value.get(u"text")
    if value_text is not None and normalize_text(value_text):
        try:
            if target_param.SetValueString(to_unicode(value_text)):
                return True
        except Exception:
            pass
        number_value = parse_number_from_text(value_text)
        if number_value is not None:
            converted_value = convert_number_for_double_target(target_param, value_text, number_value)
            if converted_value is not None:
                target_param.Set(float(converted_value))
                return True

    raw = source_value.get(u"raw")
    try:
        number_value = float(raw)
        converted_value = convert_number_for_double_target(target_param, u"", number_value)
        if converted_value is not None:
            target_param.Set(float(converted_value))
            return True
        if spec_is(target_param, u"number"):
            target_param.Set(number_value)
            return True
    except Exception:
        pass
    return False


def set_integer_value(target_param, source_value):
    source_storage = source_value.get(u"storage")
    if source_storage == DB.StorageType.Integer:
        target_param.Set(int(source_value.get(u"raw")))
        return True

    if source_storage == DB.StorageType.Double:
        target_param.Set(int(round(float(source_value.get(u"raw")))))
        return True

    number_value = parse_number_from_text(source_value.get(u"text"))
    if number_value is None:
        return False
    target_param.Set(int(round(number_value)))
    return True


def set_element_id_value(target_param, source_value):
    if source_value.get(u"storage") != DB.StorageType.ElementId:
        return False
    target_param.Set(source_value.get(u"raw"))
    return True


def write_param_value(target_param, source_value):
    if target_param is None:
        return u"parameter_not_found", u"Целевой параметр не найден"
    try:
        if target_param.IsReadOnly:
            return u"readonly", u"Целевой параметр только для чтения"
    except Exception:
        return u"exception", u"Не удалось проверить доступность параметра"

    try:
        if values_equal(target_param, source_value):
            return u"already_ok", u"Значение уже совпадает"

        st = target_param.StorageType
        if st == DB.StorageType.String:
            ok = set_string_value(target_param, source_value)
        elif st == DB.StorageType.Double:
            ok = set_double_value(target_param, source_value)
        elif st == DB.StorageType.Integer:
            ok = set_integer_value(target_param, source_value)
        elif st == DB.StorageType.ElementId:
            ok = set_element_id_value(target_param, source_value)
        else:
            return u"wrong_storage_type", u"Неподдерживаемый тип целевого параметра"

        if ok:
            return u"updated", None
        return u"conversion_failed", u"Не удалось преобразовать значение"
    except Exception as ex:
        return u"exception", to_unicode(ex)


def add_problem(problems, elem, source_choice, target_choice, reason):
    if len(problems) >= 200:
        return
    problems.append([
        get_element_id(elem),
        get_category_name(elem),
        source_choice.label,
        target_choice.label,
        reason,
    ])


def main():
    elements = get_selected_elements()
    if not elements:
        forms.alert(u"Выделите элементы перед запуском инструмента.", title=u"Копировать параметр")
        return

    choices = collect_parameters(elements)
    if not choices:
        forms.alert(u"У выделенных элементов не найдено параметров.", title=u"Копировать параметр")
        return

    source_choice = forms.SelectFromList.show(
        choices,
        name_attr="label",
        multiselect=False,
        title=u"Из какого параметра копировать"
    )
    if not source_choice:
        return

    target_choice = forms.SelectFromList.show(
        choices,
        name_attr="label",
        multiselect=False,
        title=u"В какой параметр записать"
    )
    if not target_choice:
        return

    if source_choice.scope == target_choice.scope and source_choice.name == target_choice.name:
        forms.alert(u"Источник и цель совпадают.", title=u"Копировать параметр")
        return

    stats = {
        u"processed": 0,
        u"updated": 0,
        u"already_ok": 0,
        u"problems": 0,
        u"source_empty": 0,
        u"source_missing": 0,
        u"target_missing": 0,
        u"readonly": 0,
        u"conversion_failed": 0,
        u"exception": 0,
        u"type_conflict": 0,
    }
    problems = []
    type_writes = {}

    t = DB.Transaction(doc, u"Копировать параметр")
    t.Start()
    try:
        for elem in elements:
            stats[u"processed"] += 1

            source_param, source_owner = get_param_by_choice(elem, source_choice)
            source_value = read_param_value(source_param)
            source_status = source_value.get(u"status")
            if source_status == u"missing":
                stats[u"source_missing"] += 1
                stats[u"problems"] += 1
                add_problem(problems, elem, source_choice, target_choice, u"Исходный параметр не найден")
                continue
            if source_status == u"empty":
                stats[u"source_empty"] += 1
                stats[u"problems"] += 1
                add_problem(problems, elem, source_choice, target_choice, u"Исходный параметр пуст")
                continue
            if source_status != u"ok":
                stats[u"exception"] += 1
                stats[u"problems"] += 1
                add_problem(problems, elem, source_choice, target_choice, u"Не удалось прочитать источник")
                continue

            target_param, target_owner = get_param_by_choice(elem, target_choice)
            if target_param is None:
                stats[u"target_missing"] += 1
                stats[u"problems"] += 1
                add_problem(problems, elem, source_choice, target_choice, u"Целевой параметр не найден")
                continue

            if target_choice.scope == SCOPE_TYPE:
                owner_id = get_owner_id(target_owner)
                current_key = (owner_id, target_choice.name)
                current_value = normalize_text(source_value.get(u"text"))
                if current_key in type_writes and type_writes[current_key] != current_value:
                    stats[u"type_conflict"] += 1
                    stats[u"problems"] += 1
                    add_problem(problems, elem, source_choice, target_choice, u"Один тип получает разные значения от разных экземпляров")
                    continue
                type_writes[current_key] = current_value

            status, reason = write_param_value(target_param, source_value)
            if status == u"updated":
                stats[u"updated"] += 1
            elif status == u"already_ok":
                stats[u"already_ok"] += 1
            else:
                stats[u"problems"] += 1
                if status in stats:
                    stats[status] += 1
                else:
                    stats[u"exception"] += 1
                add_problem(problems, elem, source_choice, target_choice, reason or status)

        t.Commit()
    except Exception as ex:
        try:
            if t.GetStatus() == DB.TransactionStatus.Started:
                t.RollBack()
        except Exception:
            pass
        output.print_md(u"### Ошибка")
        output.print_md(u"```\n{}\n```".format(to_unicode(ex)))
        return

    output.print_md(u"## Копирование параметра")
    output.print_md(u"> Источник: **{}** → Цель: **{}**".format(source_choice.label, target_choice.label))
    output.print_md(
        u"> Обработано: **{}** | Записано: **{}** | Уже совпадало: **{}** | Проблемы: **{}**".format(
            stats[u"processed"], stats[u"updated"], stats[u"already_ok"], stats[u"problems"]
        )
    )

    rows = [
        [u"Исходный параметр не найден", stats[u"source_missing"]],
        [u"Исходный параметр пуст", stats[u"source_empty"]],
        [u"Целевой параметр не найден", stats[u"target_missing"]],
        [u"Целевой параметр только для чтения", stats[u"readonly"]],
        [u"Не удалось преобразовать значение", stats[u"conversion_failed"]],
        [u"Конфликт записи в тип", stats[u"type_conflict"]],
        [u"Ошибки", stats[u"exception"]],
    ]
    rows = [r for r in rows if r[1] > 0]
    if rows:
        output.print_md(u"### Причины проблем")
        output.print_table(table_data=rows, columns=[u"Причина", u"Количество"], formats=[u"{}", u"{}"])

    if problems:
        table_rows = []
        for elem_id, category, source, target, reason in problems:
            try:
                id_text = output.linkify(DB.ElementId(int(elem_id)))
            except Exception:
                id_text = to_unicode(elem_id)
            table_rows.append([id_text, category, source, target, reason])
        output.print_md(u"### Проблемные элементы")
        output.print_table(
            table_data=table_rows,
            columns=[u"ID", u"Категория", u"Источник", u"Цель", u"Причина"],
            formats=[u"{}", u"{}", u"{}", u"{}", u"{}"],
        )

    forms.alert(
        u"Готово\nЗаписано: {}\nУже совпадало: {}\nПроблемы: {}".format(
            stats[u"updated"], stats[u"already_ok"], stats[u"problems"]
        ),
        title=u"Копировать параметр"
    )


if __name__ == u"__main__":
    main()
