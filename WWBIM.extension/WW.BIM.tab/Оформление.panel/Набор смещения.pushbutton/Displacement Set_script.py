
# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script

__doc__ = "Фрагмент схемы: создаёт 3D-вид узла из выбранных элементов и настраивает фильтры по параметру 'ADSK_Позиция на схеме'."

logger = script.get_logger()

UIDOC = revit.uidoc
DOC = revit.doc


def _get_param_element_id(doc, param_name):
    """Найти ElementId параметра по имени (для общего/общего проекта параметра)."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ParameterElement)
    for pe in collector:
        try:
            defn = pe.GetDefinition()
            if defn and defn.Name == param_name:
                return pe.Id
        except Exception:
            # на всякий случай гасим странные параметры
            continue
    return None


def _get_selection(uidoc):
    sel_ids = list(uidoc.Selection.GetElementIds())
    if sel_ids:
        return [DOC.GetElement(eid) for eid in sel_ids]

    forms.alert(
        u"Выберите элементы на 3D-виде и запустите команду ещё раз,\n"
        u"либо используйте окно выбора элементов.",
        title=u"Фрагмент схемы",
        warn_icon=False
    )

    try:
        picked = uidoc.Selection.PickObjects(DB.Selection.ObjectType.Element, u"Выберите элементы узла")
    except Exception:
        return []

    elems = []
    for ref in picked:
        el = DOC.GetElement(ref.ElementId)
        if el is not None:
            elems.append(el)
    return elems


def _check_active_view_is_3d(doc):
    view = doc.ActiveView
    if not isinstance(view, DB.View3D):
        forms.alert(
            u"Команда работает только на 3D-видe.\n"
            u"Переключитесь на 3D-аксонометрию и повторите.",
            title=u"Фрагмент схемы",
            warn_icon=True
        )
        return None
    if view.IsTemplate:
        forms.alert(
            u"Текущий вид является шаблоном.\n"
            u"Переключитесь на обычный 3D-вид.",
            title=u"Фрагмент схемы",
            warn_icon=True
        )
        return None
    return view


def _ask_mode():
    modes = [u"Создать новый узел", u"Добавить к существующему"]
    res = forms.CommandSwitchWindow.show(modes, message=u"Выберите действие")
    return res


def _find_node_views_for_parent(doc, parent_view):
    """Найти все виды-узлы для данного родительского вида."""
    parent_name = parent_view.Name
    all_3d = DB.FilteredElementCollector(doc).OfClass(DB.View3D)
    result = []
    for v in all_3d:
        if v.IsTemplate:
            continue
        name = v.Name
        if parent_name in name and u"Узел" in name:
            # исключаем сам родительский вид
            if v.Id != parent_view.Id:
                result.append(v)
    return result


def _ask_existing_node_view(doc, parent_view):
    node_views = _find_node_views_for_parent(doc, parent_view)
    if not node_views:
        forms.alert(
            u"Для текущего вида не найдено видов-узлов.\n"
            u"Сначала создайте узел, затем добавляйте к нему элементы.",
            title=u"Фрагмент схемы",
            warn_icon=True
        )
        return None

    # Показываем список по имени вида (свойство Name).
    chosen = forms.SelectFromList.show(
        node_views,
        title=u"Выберите существующий узел",
        button_name=u"Выбрать",
        name_attr="Name",
        width=500,
        height=400,
        multiselect=False
    )
    if not chosen:
        return None
    # При multiselect=False pyRevit возвращает сам объект вида.
    return chosen

def _get_next_node_index(doc, parent_view):
    """Определить следующий номер узла по существующим видам-узлам."""
    node_views = _find_node_views_for_parent(doc, parent_view)
    prefix = parent_view.Name + u"_Узел"
    max_index = 0
    for v in node_views:
        name = v.Name
        if not name.startswith(prefix):
            continue
        tail = name[len(prefix):]
        try:
            idx = int(tail)
            if idx > max_index:
                max_index = idx
        except Exception:
            continue
    return max_index + 1


def _ensure_param_exists(doc, elems, param_name):
    """Проверить, что у хотя бы одного элемента есть параметр с таким именем."""
    for el in elems:
        p = el.LookupParameter(param_name)
        if p is not None:
            return True
    return False


def _set_position_param(doc, elems, param_name, value):
    for el in elems:
        p = el.LookupParameter(param_name)
        if p and not p.IsReadOnly:
            try:
                p.Set(value)
            except Exception as ex:
                logger.debug(u"Не удалось записать параметр у элемента {0}: {1}".format(el.Id.IntegerValue, ex))



def _expand_with_insulations(doc, elems):
    """Добавить к списку элементов изоляцию (труб и воздуховодов),
    если она является зависимым элементом выбранных хостов."""
    if not elems:
        return elems

    all_elems = list(elems)

    pipe_ins_filter = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_PipeInsulations)
    duct_ins_filter = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_DuctInsulations)

    seen_ids = set([el.Id for el in all_elems])

    for el in elems:
        try:
            # изоляция труб
            for dep_id in el.GetDependentElements(pipe_ins_filter):
                try:
                    if dep_id not in seen_ids:
                        dep = doc.GetElement(dep_id)
                        if dep is not None:
                            all_elems.append(dep)
                            seen_ids.add(dep_id)
                except Exception:
                    continue

            # изоляция воздуховодов
            for dep_id in el.GetDependentElements(duct_ins_filter):
                try:
                    if dep_id not in seen_ids:
                        dep = doc.GetElement(dep_id)
                        if dep is not None:
                            all_elems.append(dep)
                            seen_ids.add(dep_id)
                except Exception:
                    continue
        except Exception:
            continue

    return all_elems

def _create_filters_for_node(doc, parent_view, node_view, filter_param_name, node_key, elems):
    """Создаёт два фильтра по параметру ADSK_Позиция на схеме для узла."""
    param_id = _get_param_element_id(doc, filter_param_name)
    if param_id is None:
        forms.alert(
            u"Не найден параметр проекта с именем '{0}'.\n"
            u"Проверьте, что общий параметр добавлен в проект.".format(filter_param_name),
            title=u"Фрагмент схемы",
            warn_icon=True
        )
        return

    provider = DB.ParameterValueProvider(param_id)



    # Категории фильтра: все моделируемые категории с параметром,
    # которые могут быть видимы на родительском виде.
    from System.Collections.Generic import List
    cat_ids = List[DB.ElementId]()

    # 1) выясняем, к каким категориям привязан параметр проекта
    allowed_cat_ids = set()
    try:
        param_elem = DOC.GetElement(param_id)
        if param_elem:
            defn = param_elem.GetDefinition()
            bindings = DOC.ParameterBindings
            it = bindings.ForwardIterator()
            while it.MoveNext():
                try:
                    if it.Key.Name != filter_param_name:
                        continue
                    binding = it.Current
                    catset = binding.Categories
                    if catset:
                        for cat in catset:
                            try:
                                allowed_cat_ids.add(cat.Id)
                            except Exception:
                                continue
                except Exception:
                    continue
    except Exception:
        pass

    # 2) берём только те категории модели, которые и видимы на виде, и имеют этот параметр
    all_cats = DOC.Settings.Categories
    for cat in all_cats:
        try:
            if cat.CategoryType != DB.CategoryType.Model:
                continue
            if parent_view.GetCategoryHidden(cat.Id):
                continue
            if allowed_cat_ids and cat.Id not in allowed_cat_ids:
                continue
            if cat.Id not in cat_ids:
                cat_ids.Add(cat.Id)
        except Exception:
            continue

    if cat_ids.Count == 0:
        forms.alert(
            u"Не удалось определить категории для фильтрации.\n",
            u"Проверьте, что параметр '{0}' добавлен к нужным категориям и включён на виде.".format(filter_param_name),
            title=u"Фрагмент схемы",
            warn_icon=True
        )
        return


    # Фильтр 1: элементы узла (равно node_key). Применяем к РОДИТЕЛЬСКОМУ виду и скрываем.
    equals_rule = DB.FilterStringRule(
        provider,
        DB.FilterStringEquals(),
        node_key
    )
    equals_filter = DB.ElementParameterFilter(equals_rule)

    filter1_name = node_key  # имя фильтра = имя вида узла
    try:
        pf1 = DB.ParameterFilterElement.Create(doc, filter1_name, cat_ids, equals_filter)
    except Exception:
        # возможно, фильтр уже существует — попробуем найти
        pf1 = None
        existing = DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)
        for f in existing:
            if f.Name == filter1_name:
                pf1 = f
                break
        if pf1 is None:
            raise

    if not parent_view.IsFilterApplied(pf1.Id):
        parent_view.AddFilter(pf1.Id)
    parent_view.SetFilterVisibility(pf1.Id, False)

    # Фильтр 2: все элементы, у которых значение параметра НЕ равно node_key.
    # Применяем к ВИДУ УЗЛА и скрываем, чтобы остались только элементы узла.
    # Для условия "НЕ равно" используем обратное правило от equals_rule.
    inv_rule = DB.FilterInverseRule(equals_rule)
    notequals_filter = DB.ElementParameterFilter(inv_rule)

    filter2_name = node_key + u" НЕ"
    try:
        pf2 = DB.ParameterFilterElement.Create(doc, filter2_name, cat_ids, notequals_filter)
    except Exception:
        # возможно, фильтр уже существует — попробуем найти
        pf2 = None
        existing = DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)
        for f in existing:
            if f.Name == filter2_name:
                pf2 = f
                break
        if pf2 is None:
            raise

    if not node_view.IsFilterApplied(pf2.Id):
        node_view.AddFilter(pf2.Id)
    node_view.SetFilterVisibility(pf2.Id, False)

def _fit_3d_section_to_view(doc, view3d, elems=None):
    """Подрезать 3D-вид по габаритам указанных элементов (или видимых, если elems=None)."""
    if not isinstance(view3d, DB.View3D):
        logger.debug(u"_fit_3d_section_to_view: вид не является 3D")
        return
    if view3d.IsTemplate:
        logger.debug(u"_fit_3d_section_to_view: вид является шаблоном")
        return

    # Если элементы переданы явно — используем их, иначе собираем с вида
    if elems:
        elements_to_check = elems
    else:
        collector = DB.FilteredElementCollector(doc, view3d.Id).WhereElementIsNotElementType()
        elements_to_check = list(collector)

    logger.debug(u"_fit_3d_section_to_view: проверяем {} элементов".format(len(list(elements_to_check)) if elems else "?"))

    min_pt = None
    max_pt = None

    for el in elements_to_check:
        try:
            # Используем None для получения BoundingBox в координатах модели
            bbox = el.get_BoundingBox(None)
        except Exception:
            bbox = None
        if not bbox:
            continue
        bmin = bbox.Min
        bmax = bbox.Max
        if not min_pt:
            min_pt = DB.XYZ(bmin.X, bmin.Y, bmin.Z)
            max_pt = DB.XYZ(bmax.X, bmax.Y, bmax.Z)
        else:
            min_pt = DB.XYZ(min(min_pt.X, bmin.X),
                            min(min_pt.Y, bmin.Y),
                            min(min_pt.Z, bmin.Z))
            max_pt = DB.XYZ(max(max_pt.X, bmax.X),
                            max(max_pt.Y, bmax.Y),
                            max(max_pt.Z, bmax.Z))

    if not min_pt or not max_pt:
        logger.debug(u"_fit_3d_section_to_view: не удалось получить габариты элементов")
        return

    logger.debug(u"_fit_3d_section_to_view: min=({:.2f}, {:.2f}, {:.2f}), max=({:.2f}, {:.2f}, {:.2f})".format(
        min_pt.X, min_pt.Y, min_pt.Z, max_pt.X, max_pt.Y, max_pt.Z))

    # небольшой отступ от габаритов
    dx = max_pt.X - min_pt.X
    dy = max_pt.Y - min_pt.Y
    dz = max_pt.Z - min_pt.Z
    margin = 0.1

    if dx == 0: dx = 1.0
    if dy == 0: dy = 1.0
    if dz == 0: dz = 1.0

    min_pt = DB.XYZ(min_pt.X - dx * margin,
                    min_pt.Y - dy * margin,
                    min_pt.Z - dz * margin)
    max_pt = DB.XYZ(max_pt.X + dx * margin,
                    max_pt.Y + dy * margin,
                    max_pt.Z + dz * margin)

    bbox_view = DB.BoundingBoxXYZ()
    bbox_view.Min = min_pt
    bbox_view.Max = max_pt
    bbox_view.Enabled = True

    try:
        view3d.SetSectionBox(bbox_view)
        logger.debug(u"_fit_3d_section_to_view: SectionBox установлен успешно")
    except Exception as ex:
        logger.error(u"_fit_3d_section_to_view: ошибка установки SectionBox: {}".format(ex))


def main():
    active_view = _check_active_view_is_3d(DOC)
    if active_view is None:
        return

    elems = _get_selection(UIDOC)
    elems = _expand_with_insulations(DOC, elems)
    if not elems:
        forms.alert(
            u"Не выбрано ни одного элемента.",
            title=u"Фрагмент схемы",
            warn_icon=True
        )
        return

    param_name = u"ADSK_Позиция на схеме"
    if not _ensure_param_exists(DOC, elems, param_name):
        forms.alert(
            u"У выбранных элементов нет параметра '{0}'.\n"
            u"Добавьте параметр в проект и/или категории элементов.".format(param_name),
            title=u"Фрагмент схемы",
            warn_icon=True
        )
        return

    mode = _ask_mode()
    if not mode:
        return

    t = DB.Transaction(DOC, u"Фрагмент схемы")
    t.Start()

    try:
        if mode == u"Создать новый узел":
            # определяем номер узла и создаём новый вид
            next_index = _get_next_node_index(DOC, active_view)
            node_view_name = u"{0}_Узел{1}".format(active_view.Name, next_index)

            dup_id = active_view.Duplicate(DB.ViewDuplicateOption.Duplicate)
            node_view = DOC.GetElement(dup_id)
            node_view.Name = node_view_name

            # записываем значение параметра у выбранных элементов
            _set_position_param(DOC, elems, param_name, node_view_name)

            # создаём/настраиваем фильтры
            _create_filters_for_node(DOC, active_view, node_view, param_name, node_view_name, elems)
            _fit_3d_section_to_view(DOC, node_view, elems)

            msg = u"Создан узел '{0}' и настроены фильтры.\nЭлементы перенесены на новый вид.".format(node_view_name)
            logger.info(msg)

        elif mode == u"Добавить к существующему":
            node_view = _ask_existing_node_view(DOC, active_view)
            if node_view is None:
                t.RollBack()
                return

            node_key = node_view.Name

            # проверяем, что для выбранного вида уже настроены фильтры (мягкая проверка)
            # если фильтров нет — просто изменим параметр у элементов, а пользователь настроит фильтры сам.
            _set_position_param(DOC, elems, param_name, node_key)
            _create_filters_for_node(DOC, active_view, node_view, param_name, node_key, elems)
            _fit_3d_section_to_view(DOC, node_view, elems)

            msg = u"Элементы добавлены к узлу '{0}'.\n"
            msg += u"Если фильтры уже были настроены, элементы скроются на общем виде и появятся на виде узла."
            logger.info(msg.format(node_key))

        t.Commit()

    except Exception as exc:
        logger.error(u"Ошибка при выполнении команды: {0}".format(exc))
        try:
            t.RollBack()
        except Exception:
            pass
        forms.alert(
            u"Произошла ошибка при выполнении команды.\nПодробности смотрите в журнале pyRevit.",
            title=u"Фрагмент схемы",
            warn_icon=True
        )


if __name__ == '__main__':
    main()