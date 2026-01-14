# -*- coding: utf-8 -*-
"""
Синхронизация параметров кубиков из связи с "Отверстия" в названии.
Сопоставляет кубики по параметру "№ кубика" и обновляет параметры FutureBIM.
"""
from __future__ import print_function, division

import sys
from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# Параметры FutureBIM для синхронизации
MATCH_PARAM = u"№ кубика"  # параметр для сопоставления

SYNC_PARAMS = [
    u"Статус",
    u"Согласовано АР",
    u"Согласовано КР",
    u"Согласовано ИОС",
    u"Комментарий АР",
    u"Комментарий КР",
    u"Комментарий ИОС",
    u"КОМ_АР",
    u"КОМ_КР",
    u"КОМ_ИОС",
    u"Тип системы",
]


def format_value(val):
    """Форматировать значение для отображения."""
    if val is None:
        return u"<пусто>"
    if isinstance(val, bool):
        return u"Да" if val else u"Нет"
    if isinstance(val, int):
        # Для булевых параметров (0/1)
        if val == 0:
            return u"0 (Нет)"
        elif val == 1:
            return u"1 (Да)"
        return unicode(val)
    if isinstance(val, float):
        return u"{:.2f}".format(val)
    if val == u"" or val == "":
        return u"<пусто>"
    return unicode(val)


def get_param_value(elem, param_name):
    """Получить значение параметра."""
    p = elem.LookupParameter(param_name)
    if p is None:
        return None

    st = p.StorageType
    if st == DB.StorageType.String:
        return p.AsString()
    elif st == DB.StorageType.Integer:
        return p.AsInteger()
    elif st == DB.StorageType.Double:
        return p.AsDouble()
    elif st == DB.StorageType.ElementId:
        return p.AsElementId().IntegerValue
    return None


def get_param_value_string(elem, param_name):
    """Получить значение параметра как отображаемую строку (AsValueString)."""
    p = elem.LookupParameter(param_name)
    if p is None:
        return None
    try:
        vs = p.AsValueString()
        if vs:
            return vs
    except:
        pass
    return get_param_value(elem, param_name)


def set_param_value(elem, param_name, value):
    """Установить значение параметра."""
    p = elem.LookupParameter(param_name)
    if p is None or p.IsReadOnly:
        return False

    if value is None:
        return False

    st = p.StorageType
    try:
        if st == DB.StorageType.String:
            p.Set(unicode(value) if value is not None else u"")
        elif st == DB.StorageType.Integer:
            p.Set(int(value))
        elif st == DB.StorageType.Double:
            p.Set(float(value))
        elif st == DB.StorageType.ElementId:
            p.Set(DB.ElementId(int(value)))
        return True
    except:
        return False


def find_link_with_openings(doc):
    """Найти связь с 'Отверстия' в названии."""
    links = DB.FilteredElementCollector(doc)\
        .OfClass(DB.RevitLinkInstance)\
        .ToElements()

    matching_links = []
    for link in links:
        link_name = link.Name
        if u"Отверстия" in link_name or u"отверстия" in link_name.lower():
            matching_links.append(link)

    return matching_links


def get_generic_models_from_link(link_instance):
    """Получить все обобщённые модели из связи."""
    link_doc = link_instance.GetLinkDocument()
    if link_doc is None:
        return []

    collector = DB.FilteredElementCollector(link_doc)\
        .OfCategory(DB.BuiltInCategory.OST_GenericModel)\
        .WhereElementIsNotElementType()\
        .ToElements()

    return list(collector)


def get_generic_models_from_doc(doc):
    """Получить все обобщённые модели из документа."""
    collector = DB.FilteredElementCollector(doc)\
        .OfCategory(DB.BuiltInCategory.OST_GenericModel)\
        .WhereElementIsNotElementType()\
        .ToElements()

    return list(collector)


def build_kubik_dict(elements, param_name):
    """Построить словарь {номер_кубика: элемент}."""
    result = {}
    for elem in elements:
        val = get_param_value(elem, param_name)
        if val is not None and val != u"" and val != "":
            key = unicode(val).strip()
            if key:
                result[key] = elem
    return result


def sync_parameters(source_elem, target_elem, param_names):
    """Синхронизировать параметры от источника к цели.
    Возвращает список кортежей: (param_name, old_value, new_value, old_display, new_display)
    """
    synced = []
    for pname in param_names:
        src_val = get_param_value(source_elem, pname)
        if src_val is not None:
            tgt_val = get_param_value(target_elem, pname)
            # Синхронизируем только если значения отличаются
            if src_val != tgt_val:
                # Получаем отображаемые значения для отчёта
                src_display = get_param_value_string(source_elem, pname)
                tgt_display = get_param_value_string(target_elem, pname)
                if set_param_value(target_elem, pname, src_val):
                    synced.append((pname, tgt_val, src_val, tgt_display, src_display))
    return synced


def main():
    # 1. Найти связи с "Отверстия"
    links = find_link_with_openings(doc)

    if not links:
        forms.alert(
            u"Не найдена связь с 'Отверстия' в названии.\n\n"
            u"Убедитесь, что связь с заданиями на отверстия загружена.",
            title=u"Синхронизация кубиков",
            warn_icon=True
        )
        return

    # Если несколько связей - дать выбрать
    if len(links) > 1:
        class LinkItem:
            def __init__(self, link):
                self.link = link
                self.name = link.Name
            def __str__(self):
                return self.name

        items = [LinkItem(l) for l in links]
        selected = forms.SelectFromList.show(
            items,
            title=u"Выберите связь с отверстиями",
            multiselect=False
        )
        if not selected:
            return
        link = selected.link
    else:
        link = links[0]

    link_doc = link.GetLinkDocument()
    if link_doc is None:
        forms.alert(
            u"Не удалось открыть документ связи.\n"
            u"Возможно, связь выгружена.",
            title=u"Синхронизация кубиков",
            warn_icon=True
        )
        return

    # 2. Получить кубики из связи и текущего документа
    link_kubiks = get_generic_models_from_link(link)
    doc_kubiks = get_generic_models_from_doc(doc)

    if not link_kubiks:
        forms.alert(
            u"В связи '{}' не найдены обобщённые модели (кубики).".format(link.Name),
            title=u"Синхронизация кубиков",
            warn_icon=True
        )
        return

    if not doc_kubiks:
        forms.alert(
            u"В текущем документе не найдены обобщённые модели (кубики).",
            title=u"Синхронизация кубиков",
            warn_icon=True
        )
        return

    # 3. Построить словари по номеру кубика
    link_dict = build_kubik_dict(link_kubiks, MATCH_PARAM)
    doc_dict = build_kubik_dict(doc_kubiks, MATCH_PARAM)

    if not link_dict:
        forms.alert(
            u"В связи не найдены кубики с заполненным параметром '{}'.".format(MATCH_PARAM),
            title=u"Синхронизация кубиков",
            warn_icon=True
        )
        return

    if not doc_dict:
        forms.alert(
            u"В документе не найдены кубики с заполненным параметром '{}'.".format(MATCH_PARAM),
            title=u"Синхронизация кубиков",
            warn_icon=True
        )
        return

    # 4. Найти совпадения и синхронизировать
    matched = 0
    updated = 0
    params_updated = 0
    not_found_in_link = []

    # Детальная информация об изменениях
    # {kubik_num: [(param_name, old_val, new_val, old_display, new_display), ...]}
    changes_log = {}

    t = DB.Transaction(doc, u"Синхронизация параметров кубиков")
    t.Start()

    try:
        for kubik_num, doc_elem in doc_dict.items():
            if kubik_num in link_dict:
                matched += 1
                link_elem = link_dict[kubik_num]
                synced = sync_parameters(link_elem, doc_elem, SYNC_PARAMS)
                if synced:
                    updated += 1
                    params_updated += len(synced)
                    changes_log[kubik_num] = {
                        'elem_id': doc_elem.Id.IntegerValue,
                        'changes': synced
                    }
            else:
                not_found_in_link.append(kubik_num)

        t.Commit()
    except Exception as e:
        t.RollBack()
        forms.alert(
            u"Ошибка при синхронизации: {}".format(str(e)),
            title=u"Синхронизация кубиков",
            warn_icon=True
        )
        return

    # 5. Детальный отчёт в окне вывода pyRevit
    output.print_md(u"# Синхронизация кубиков")
    output.print_md(u"**Связь:** {}".format(link.Name))
    output.print_md(u"---")

    output.print_md(u"## Статистика")
    output.print_md(u"| Показатель | Значение |")
    output.print_md(u"|------------|----------|")
    output.print_md(u"| Кубиков в связи | {} |".format(len(link_dict)))
    output.print_md(u"| Кубиков в документе | {} |".format(len(doc_dict)))
    output.print_md(u"| Сопоставлено | {} |".format(matched))
    output.print_md(u"| Обновлено кубиков | {} |".format(updated))
    output.print_md(u"| Обновлено параметров | {} |".format(params_updated))

    # Детали изменений
    if changes_log:
        output.print_md(u"---")
        output.print_md(u"## Детали изменений")
        output.print_md(u"")

        for kubik_num, data in sorted(changes_log.items()):
            elem_id = data['elem_id']
            changes = data['changes']

            output.print_md(u"### Кубик: {} (ID: {})".format(kubik_num, elem_id))
            output.print_md(u"| Параметр | Было | Стало |")
            output.print_md(u"|----------|------|-------|")

            for change in changes:
                pname, old_val, new_val, old_display, new_display = change
                old_str = format_value(old_display if old_display else old_val)
                new_str = format_value(new_display if new_display else new_val)
                output.print_md(u"| {} | {} | {} |".format(pname, old_str, new_str))

            output.print_md(u"")
    else:
        output.print_md(u"---")
        output.print_md(u"*Изменений не обнаружено - все параметры уже синхронизированы.*")

    # Не найденные в связи
    if not_found_in_link:
        output.print_md(u"---")
        output.print_md(u"## Не найдены в связи ({})".format(len(not_found_in_link)))
        output.print_md(u"Эти кубики есть в документе, но отсутствуют в связи:")
        output.print_md(u"")
        for num in sorted(not_found_in_link)[:20]:
            output.print_md(u"- {}".format(num))
        if len(not_found_in_link) > 20:
            output.print_md(u"- *... и ещё {}*".format(len(not_found_in_link) - 20))

    output.print_md(u"---")
    output.print_md(u"*Синхронизация завершена.*")


if __name__ == "__main__":
    main()
