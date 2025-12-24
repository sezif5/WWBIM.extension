# -*- coding: utf-8 -*-
# pyRevit button script: 3D-вид по ID элемента в выбранной загруженной связи
# - Всегда показывает список только ЗАГРУЖЕННЫХ связей и просит выбрать,
#   в какой связи искать элемент по ID.
# - Использует уже открытый 3D-вид, если он активен,
#   иначе 3D-вид пользователя вида {3D - Username},
#   и только если их нет — создаёт новый 3D-вид.
# - В Revit 2023+ пытается подсветить КОНКРЕТНЫЙ элемент в связи через
#   Selection.SetReferences + Reference.CreateLinkReference.
#   В более старых версиях API выделяется только экземпляр связи.

import clr
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ElementId,
    View3D,
    ViewFamily,
    ViewFamilyType,
    Transaction,
    BoundingBoxXYZ,
    XYZ,
    Reference,
)
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import Clipboard
from System.Collections.Generic import List as CsList

from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


def _show(title, text):
    try:
        TaskDialog.Show(title, text)
    except:
        pass


def _get_id_from_user():
    """Получить целочисленный ID элемента.
    Пытаемся взять текст из буфера обмена как значение по умолчанию,
    затем просим пользователя подтвердить/изменить его."""
    clip_text = None
    try:
        if Clipboard.ContainsText():
            clip_text = Clipboard.GetText()
    except:
        clip_text = None

    default = u""
    if clip_text:
        default = clip_text.strip()

    id_text = None
    try:
        id_text = forms.ask_for_string(
            prompt=u"Введите ID элемента из связанного файла (можно просто нажать Ctrl+V)",
            default=default,
            title=u"3D по ID элемента в связи",
        )
    except:
        _show(u"Ввод ID", u"Не удалось открыть форму ввода ID.")
        return None

    if id_text is None:
        return None

    id_text = id_text.strip()
    if not id_text:
        return None

    try:
        return int(id_text)
    except:
        _show(u"Некорректный ID", u"Не удалось преобразовать '{0}' в целое число.".format(id_text))
        return None


def _choose_link_instance():
    """Показать список всех ЗАГРУЖЕННЫХ связей и вернуть выбранный RevitLinkInstance."""
    links = []
    labels = []

    col = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    for link_inst in col:
        # Используем только загруженные связи (у которых есть LinkDocument)
        linked_doc = None
        try:
            linked_doc = link_inst.GetLinkDocument()
        except:
            linked_doc = None

        if linked_doc is None:
            continue

        links.append(link_inst)

        try:
            link_name = link_inst.Name
        except:
            link_name = u"(без имени)"

        try:
            doc_title = linked_doc.Title
        except:
            doc_title = u""

        label = u"{0} | {1}".format(link_name, doc_title)
        labels.append(label)

    if not links:
        _show(u"Связи", u"В документе нет загруженных связей.")
        return None

    if len(links) == 1:
        # одна загруженная связь — выбираем её автоматически
        return links[0]

    choice = forms.SelectFromList.show(
        labels,
        title=u"Выбор связи для поиска ID",
        button_name=u"Искать в этой связи",
        multiselect=False,
    )

    if not choice:
        return None

    idx = labels.index(choice)
    return links[idx]


def _find_element_in_links(element_int_id):
    """Найти элемент с указанным ID в выбранной пользователем связи.
    Возвращает (RevitLinkInstance, Element) или (None, None)."""
    target_id = ElementId(element_int_id)

    link_inst = _choose_link_instance()
    if link_inst is None:
        return (None, None)

    linked_doc = None
    try:
        linked_doc = link_inst.GetLinkDocument()
    except:
        linked_doc = None

    if linked_doc is None:
        _show(u"Связь", u"Не удалось получить документ выбранной связи.")
        return (None, None)

    try:
        linked_el = linked_doc.GetElement(target_id)
    except:
        linked_el = None

    if linked_el is None:
        _show(
            u"Поиск элемента",
            u"Элемент с ID {0} не найден в выбранной связи.".format(element_int_id),
        )
        return (None, None)

    return (link_inst, linked_el)


def _get_3d_view_family_type_id():
    vft_col = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    for vft in vft_col:
        try:
            if vft.ViewFamily == ViewFamily.ThreeDimensional:
                return vft.Id
        except:
            continue
    return None


def _get_existing_or_personal_3d_view():
    """Вернуть уже открытый 3D-вид (если активен) или персональный {3D - Username}.
    Если ни один не найден, вернуть None (в этом случае создадим новый)."""
    # 1) Если активный вид — 3D и не шаблон, используем его
    av = uidoc.ActiveView
    try:
        if isinstance(av, View3D) and not av.IsTemplate:
            return av
    except:
        pass

    # 2) Ищем персональный 3D-вид пользователя вида {3D - Username}
    try:
        username = doc.Application.Username
    except:
        username = None

    if username:
        target_name = u"{3D - %s}" % username
        v_col = FilteredElementCollector(doc).OfClass(View3D)
        for v in v_col:
            try:
                if (not v.IsTemplate) and v.Name == target_name:
                    return v
            except:
                continue

    return None


def _prepare_3d_view_with_section_box(link_inst, linked_el, element_int_id):
    """Подготовить 3D-вид и обрезать его по элементу в связи.
    Использует существующий 3D-вид или персональный, а при их отсутствии создаёт новый.
    Возвращает View3D либо None."""
    # Границы элемента в связанном файле
    try:
        bb_linked = linked_el.get_BoundingBox(None)
    except:
        bb_linked = None

    if bb_linked is None:
        _show(u"3D вид", u"Не удалось получить границы элемента в связи.")
        return None

    try:
        transform = link_inst.GetTotalTransform()
    except:
        transform = None

    if transform is None:
        _show(u"3D вид", u"Не удалось получить трансформацию связи.")
        return None

    min_pt_link = bb_linked.Min
    max_pt_link = bb_linked.Max

    # преобразуем в координаты хост-документа
    min_host = transform.OfPoint(min_pt_link)
    max_host = transform.OfPoint(max_pt_link)

    min_x = min(min_host.X, max_host.X)
    min_y = min(min_host.Y, max_host.Y)
    min_z = min(min_host.Z, max_host.Z)

    max_x = max(min_host.X, max_host.X)
    max_y = max(min_host.Y, max_host.Y)
    max_z = max(min_host.Z, max_host.Z)

    # небольшой отступ вокруг элемента (в футах)
    pad = 3.0
    min_xyz = XYZ(min_x - pad, min_y - pad, min_z - pad)
    max_xyz = XYZ(max_x + pad, max_y + pad, max_z + pad)

    view3d = None

    t = Transaction(doc, u"3D по ID элемента в связи")
    t.Start()
    try:
        # 1) Пытаемся использовать уже существующий 3D-вид
        view3d = _get_existing_or_personal_3d_view()

        # 2) Если не нашли — создаём новый
        if view3d is None:
            vft_id = _get_3d_view_family_type_id()
            if vft_id is None:
                t.RollBack()
                _show(u"3D вид", u"В документе не найден тип 3D-вида.")
                return None
            view3d = View3D.CreateIsometric(doc, vft_id)

        # Обрезка секционным боксом
        bbox = BoundingBoxXYZ()
        bbox.Min = min_xyz
        bbox.Max = max_xyz

        try:
            view3d.SetSectionBox(bbox)
        except:
            view3d.SectionBox = bbox

        try:
            view3d.IsSectionBoxActive = True
        except:
            pass

        t.Commit()
    except:
        t.RollBack()
        raise

    return view3d


def _select_link_or_element_in_view(view3d, link_inst, linked_el):
    """Делаем 3D-вид активным и по возможности выделяем КОНКРЕТНЫЙ элемент
    в связи (через Selection.SetReferences + Reference.CreateLinkReference).
    Если API этого не поддерживает, просто выделяем экземпляр связи."""
    if view3d is None:
        return

    try:
        uidoc.ActiveView = view3d
    except:
        return

    selection = uidoc.Selection

    # Пытаемся использовать Selection.SetReferences (Revit 2023+)
    has_set_refs = False
    try:
        # hasattr на .NET-метод в IronPython работает
        has_set_refs = hasattr(selection, 'SetReferences')
    except:
        has_set_refs = False

    if has_set_refs:
        try:
            # Reference на элемент в связке
            ref_in_link = Reference(linked_el)
            # Преобразуем в reference в хосте для конкретного экземпляра связи
            ref_in_host = ref_in_link.CreateLinkReference(link_inst)

            from System.Collections.Generic import List as CsRefList
            ref_list = CsRefList[Reference]()
            ref_list.Add(ref_in_host)

            selection.SetReferences(ref_list)
            uidoc.ShowElements(ref_in_host)
            return
        except:
            # Если что-то пошло не так — падаем в запасной вариант
            pass

    # Fallback: просто выделяем экземпляр связи (как минимум не падаем)
    try:
        ids = CsList[ElementId]()
        ids.Add(link_inst.Id)
        selection.SetElementIds(ids)
        uidoc.ShowElements(link_inst.Id)
    except:
        pass


def main():
    element_int_id = _get_id_from_user()
    if element_int_id is None:
        return

    link_inst, linked_el = _find_element_in_links(element_int_id)
    if link_inst is None or linked_el is None:
        return

    view3d = _prepare_3d_view_with_section_box(link_inst, linked_el, element_int_id)
    _select_link_or_element_in_view(view3d, link_inst, linked_el)


if __name__ == "__main__":
    main()
