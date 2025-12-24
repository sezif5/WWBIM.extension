# -*- coding: utf-8 -*-

import clr

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSchedule,
    ElementId,
    ElementIdSetFilter,
    ScheduleSheetInstance,
    ViewSheet,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

from pyrevit import script, forms

from System.Collections.Generic import List as Clist

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
output = script.get_output()


def get_target_element():
    """Возвращает элемент: либо заранее выбранный, либо выбранный пользователем."""
    sel_ids = uidoc.Selection.GetElementIds()
    if sel_ids and sel_ids.Count == 1:
        elem_id = list(sel_ids)[0]
        elem = doc.GetElement(elem_id)
        if elem is not None:
            return elem

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Выберите элемент для проверки в спецификациях"
        )
        if ref:
            return doc.GetElement(ref.ElementId)
    except:
        return None

    return None


def build_schedule_sheet_map():
    """Строит словарь {ScheduleId.IntegerValue: [ViewSheet, ...]} для всех размещённых спецификаций."""
    result = {}
    try:
        ssi_collector = FilteredElementCollector(doc).OfClass(ScheduleSheetInstance)
    except:
        return result

    for inst in ssi_collector:
        try:
            sch_id = inst.ScheduleId
            sheet_id = inst.OwnerViewId
            sheet = doc.GetElement(sheet_id)
            if sheet is None:
                continue
            key = sch_id.IntegerValue
            sheets_list = result.get(key)
            if sheets_list is None:
                sheets_list = []
                result[key] = sheets_list
            sheets_list.append(sheet)
        except:
            continue

    return result


def find_schedules_for_element(element, only_on_sheets=False, schedule_sheet_map=None):
    """Возвращает список спецификаций, в которых присутствует переданный элемент.

    Оптимизации:
    - фильтр по Id элемента (ElementIdSetFilter);
    - отбор спецификаций по CategoryId (Definition.CategoryId),
      при этом многоцелевые (multi-category) спецификации не отбрасываются;
    - прогресс-бар pyRevit с возможностью отмены.
    """
    if element is None:
        return []

    element_id = element.Id
    category = element.Category
    cat_id = category.Id if category else None

    # фильтр по конкретному Id элемента
    id_list = Clist[ElementId]()
    id_list.Add(element_id)
    id_filter = ElementIdSetFilter(id_list)

    result_schedules = []

    schedules = list(FilteredElementCollector(doc).OfClass(ViewSchedule))
    
    # Если нужно проверять только спецификации на листах
    if only_on_sheets and schedule_sheet_map:
        schedules = [s for s in schedules if s.Id.IntegerValue in schedule_sheet_map]
    
    total = len(schedules)
    if total == 0:
        return result_schedules

    # прогресс-бар с возможностью отмены
    with forms.ProgressBar(
        title=u"Проверка спецификаций: {value} из {max_value}",
        cancellable=True,
        step=5
    ) as pb:
        for idx, schedule in enumerate(schedules):
            # обновляем прогресс
            try:
                pb.update_progress(idx + 1, total)
            except:
                pass

            # проверка на отмену пользователем
            try:
                if getattr(pb, 'cancelled', False):
                    break
            except:
                pass

            if schedule is None:
                continue

            # пропускаем шаблоны спецификаций
            try:
                if schedule.IsTemplate:
                    continue
            except:
                pass

            # быстрый отбор по основной категории спецификации
            try:
                defn = schedule.Definition
                sch_cat_id = defn.CategoryId if defn else ElementId.InvalidElementId
            except:
                sch_cat_id = ElementId.InvalidElementId

            # если у спецификации задана одна категория и она не совпадает с категорией элемента — пропускаем
            if cat_id is not None                     and sch_cat_id is not None                     and sch_cat_id != ElementId.InvalidElementId                     and sch_cat_id != cat_id:
                continue

            # коллектор в контексте спецификации
            try:
                collector = FilteredElementCollector(doc, schedule.Id)
            except:
                continue

            # дополнительный фильтр по категории — где применимо
            if cat_id is not None:
                try:
                    collector = collector.OfCategoryId(cat_id)
                except:
                    pass

            # фильтр по Id
            try:
                collector = collector.WherePasses(id_filter)
            except:
                continue

            # проверяем, есть ли хотя бы один элемент
            found = False
            try:
                if collector.GetElementCount() > 0:
                    found = True
            except:
                pass

            if not found:
                # fallback через итератор
                try:
                    it = collector.GetElementIterator()
                    it.Reset()
                    if it.MoveNext():
                        found = True
                except:
                    pass

            if found:
                result_schedules.append(schedule)

    return result_schedules


def show_schedules_table(schedules, schedule_sheet_map):
    """Показывает список спецификаций в красивом формате с эмодзи и цветами."""
    if not schedules:
        output.print_md(u"---")
        output.print_md(u"## ❌ Результат проверки")
        output.print_md(u"")
        output.print_md(u"<span style='color:#e74c3c; font-size:14px;'>**Элемент не найден ни в одной спецификации проекта.**</span>")
        output.print_md(u"")
        return

    # Заголовок
    output.print_md(u"---")
    output.print_md(u"## ✅ Результат проверки")
    output.print_md(u"")
    output.print_md(u"<span style='color:#27ae60; font-size:14px;'>**Элемент найден в {} спецификациях:**</span>".format(len(schedules)))
    output.print_md(u"")
    
    # Разделяем на размещённые на листах и не размещённые
    on_sheets = []
    not_on_sheets = []
    
    for sch in sorted(schedules, key=lambda x: x.Name):
        sheets = schedule_sheet_map.get(sch.Id.IntegerValue)
        if sheets:
            on_sheets.append((sch, sheets))
        else:
            not_on_sheets.append(sch)
    
    # Спецификации на листах
    if on_sheets:
        output.print_md(u"### 📋 Спецификации на листах")
        output.print_md(u"")
        
        table_data = []
        for idx, (sch, sheets) in enumerate(on_sheets, 1):
            spec_link = output.linkify(sch.Id, title=sch.Name)
            
            sheet_links = []
            for sh in sheets:
                try:
                    title = u"{}  {}".format(sh.SheetNumber, sh.Name)
                except:
                    title = sh.Name
                sheet_links.append(output.linkify(sh.Id, title=title))
            sheet_cell = u"<br>".join(sheet_links)
            
            table_data.append([idx, spec_link, sheet_cell])
        
        output.print_table(table_data, columns=[u"№", u"📑 Спецификация", u"📄 Лист"])
        output.print_md(u"")
    
    # Спецификации НЕ на листах
    if not_on_sheets:
        output.print_md(u"### 📝 Спецификации не размещены на листах")
        output.print_md(u"")
        
        table_data = []
        for idx, sch in enumerate(not_on_sheets, 1):
            spec_link = output.linkify(sch.Id, title=sch.Name)
            table_data.append([idx, spec_link, u"⚠️ Не на листе"])
        
        output.print_table(table_data, columns=[u"№", u"📑 Спецификация", u"Статус"])
        output.print_md(u"")
    
    # Итог
    output.print_md(u"---")
    output.print_md(u"### 📊 Итого")
    output.print_md(u"")
    
    summary_data = [
        [u"✅ Всего спецификаций", len(schedules)],
        [u"📋 На листах", len(on_sheets)],
        [u"⚠️ Не размещены", len(not_on_sheets)],
    ]
    output.print_table(summary_data, columns=[u"Показатель", u"Количество"])
    output.print_md(u"")


def main():
    element = get_target_element()
    if element is None:
        TaskDialog.Show(
            "Проверка спецификаций",
            "Не удалось получить элемент. Выберите один элемент и запустите скрипт снова."
        )
        return

    schedule_sheet_map = build_schedule_sheet_map()
    
    # Проверяем количество спецификаций
    all_schedules = list(FilteredElementCollector(doc).OfClass(ViewSchedule))
    only_on_sheets = False
    
    if len(all_schedules) > 50:
        result = forms.alert(
            u"Спецификаций для проверки > 50. Проверять только спецификации, размещённые на листах?",
            yes=True,
            no=True,
            cancel=True
        )
        if result is None:  # Cancel
            return
        only_on_sheets = result  # True = да, False = нет
    
    schedules = find_schedules_for_element(element, only_on_sheets, schedule_sheet_map)
    show_schedules_table(schedules, schedule_sheet_map)


if __name__ == "__main__":
    main()
