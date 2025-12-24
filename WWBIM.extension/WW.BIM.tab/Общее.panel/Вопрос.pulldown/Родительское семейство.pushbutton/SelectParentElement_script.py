# -*- coding: utf-8 -*-
# pyRevit script: Select parent element for nested element
# Скрипт выделяет родительский элемент для выбранного вложенного элемента.
# Если до запуска был выбран элемент, используется он, иначе — будет предложено выбрать элемент в модели.

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def get_initial_element_id():
    sel_ids = uidoc.Selection.GetElementIds()
    if sel_ids and sel_ids.Count > 0:
        # Берём первый выбранный элемент
        return list(sel_ids)[0]
    # Если предварительного выбора нет — просим выбрать элемент
    ref = uidoc.Selection.PickObject(ObjectType.Element, "Выберите вложенный (дочерний) элемент")
    return ref.ElementId

def find_parent_element(elem):
    # 1. SuperComponent — для вложенных/массивов/подкомпонентов
    try:
        super_comp = getattr(elem, "SuperComponent", None)
    except Exception:
        super_comp = None

    if super_comp is not None:
        return super_comp

    # 2. Host для семейств, у которых есть хост (двери, окна и т.п.)
    fi = elem if isinstance(elem, FamilyInstance) else None
    if fi is not None:
        host = fi.Host
        if host is not None:
            return host

    # 3. AssemblyInstanceId — если элемент входит в сборку
    try:
        asm_id = elem.AssemblyInstanceId
    except Exception:
        asm_id = ElementId.InvalidElementId

    if asm_id is not None and asm_id != ElementId.InvalidElementId:
        asm = doc.GetElement(asm_id)
        if asm is not None:
            return asm

    # 4. GroupId — если элемент входит в группу
    try:
        group_id = elem.GroupId
    except Exception:
        group_id = ElementId.InvalidElementId

    if group_id is not None and group_id != ElementId.InvalidElementId:
        grp = doc.GetElement(group_id)
        if grp is not None:
            return grp

    # Родитель не найден
    return None

def select_element(elem):
    ids = List[ElementId]()
    ids.Add(elem.Id)
    uidoc.Selection.SetElementIds(ids)

def main():
    try:
        elem_id = get_initial_element_id()
    except Exception:
        TaskDialog.Show("PyRevit", "Не удалось получить выбранный элемент.")
        return

    elem = doc.GetElement(elem_id)
    if elem is None:
        TaskDialog.Show("PyRevit", "Не удалось получить элемент по идентификатору.")
        return

    parent = find_parent_element(elem)
    if parent is None:
        TaskDialog.Show("PyRevit", "Родительский элемент для выбранного элемента не найден.")
        return

    select_element(parent)

if __name__ == "__main__":
    main()
