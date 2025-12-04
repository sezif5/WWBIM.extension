# -*- coding: utf-8 -*-
from __future__ import division

# PyRevit / Revit API imports
from pyrevit import revit, DB, UI, script, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = revit.doc
uidoc = revit.uidoc


def has_connectors(element):
    """Return True if element has at least one MEP connector."""
    cm = None
    # FamilyInstance with MEPModel
    mep_model = getattr(element, "MEPModel", None)
    if mep_model:
        try:
            cm = mep_model.ConnectorManager
        except:
            cm = None
    # MEPCurve (Pipe, Duct, CableTray, etc.)
    if cm is None and hasattr(element, "ConnectorManager"):
        try:
            cm = element.ConnectorManager
        except:
            cm = None

    if not cm:
        return False

    try:
        for c in cm.Connectors:
            # as soon as we find one connector, we know it's connectable
            return True
    except:
        return False
    return False


def get_connectors(element):
    """Get list of Autodesk.Revit.DB.Connector from element."""
    connectors = []
    cm = None
    mep_model = getattr(element, "MEPModel", None)
    if mep_model:
        try:
            cm = mep_model.ConnectorManager
        except:
            cm = None
    if cm is None and hasattr(element, "ConnectorManager"):
        try:
            cm = element.ConnectorManager
        except:
            cm = None
    if not cm:
        return connectors
    try:
        for c in cm.Connectors:
            connectors.append(c)
    except:
        pass
    return connectors


class ConnectorSelectionFilter(ISelectionFilter):
    """Selection filter that allows only elements with MEP connectors."""
    def AllowElement(self, element):
        return has_connectors(element)

    def AllowReference(self, reference, point):
        return True


class ConnectorOption(object):
    """Wrapper for displaying connector choices in pyRevit SelectFromList."""
    def __init__(self, owner, connector, index):
        self.owner = owner
        self.connector = connector
        self.index = index
        self.name = self._build_name()

    def _build_name(self):
        try:
            owner_name = self.owner.Name
            if not owner_name:
                owner_name = u"Id {0}".format(self.owner.Id.IntegerValue)
        except:
            owner_name = u"Id {0}".format(self.owner.Id.IntegerValue)

        origin = self.connector.Origin
        domain = self.connector.Domain.ToString()
        status = self.connector.IsConnected and u"занят" or u"свободен"

        size_info = u""
        try:
            # many MEP connectors expose Radius in feet
            radius = getattr(self.connector, "Radius", None)
            if radius:
                diam_mm = 2.0 * radius * 304.8
                size_info = u" D≈{0:.0f}мм".format(diam_mm)
        except:
            size_info = u""

        try:
            x_mm = origin.X * 304.8
            y_mm = origin.Y * 304.8
            z_mm = origin.Z * 304.8
        except:
            x_mm = y_mm = z_mm = 0.0

        return u"#{0} | {1} | {2} | ({3:.0f}; {4:.0f}; {5:.0f}){6}".format(
            self.index + 1,
            owner_name,
            domain,
            x_mm, y_mm, z_mm,
            u" | " + status + size_info
        )


def choose_connector(element, title):
    """Ask user to choose free connector from element.

    Показывает только свободные коннекторы.
    Если свободный всего один, выбирает его автоматически.
    """
    all_connectors = get_connectors(element)
    if not all_connectors:
        forms.alert(
            u"У выбранного элемента (Id {0}) нет MEP-коннекторов.".format(
                element.Id.IntegerValue
            ),
            ok=True,
            exitscript=True
        )

    free_connectors = []
    for c in all_connectors:
        try:
            if not c.IsConnected:
                free_connectors.append(c)
        except:
            # если по какой-то причине IsConnected недоступен,
            # считаем коннектор свободным
            free_connectors.append(c)

    if not free_connectors:
        forms.alert(
            u"У элемента (Id {0}) нет свободных коннекторов.".format(
                element.Id.IntegerValue
            ),
            ok=True,
            exitscript=True
        )

    # если всего один свободный коннектор — выбираем его без диалога
    if len(free_connectors) == 1:
        return free_connectors[0]

    options = []
    for idx, conn in enumerate(free_connectors):
        options.append(ConnectorOption(element, conn, idx))

    chosen = forms.SelectFromList.show(
        options,
        title=title,
        multiselect=False,
        name_attr="name"
    )

    if not chosen:
        script.exit()

    return chosen.connector


def find_closest_connector(element, point):
    """Find connector on element whose origin is closest to given XYZ."""
    connectors = get_connectors(element)
    if not connectors:
        return None
    nearest = None
    min_dist = None
    for c in connectors:
        try:
            d = c.Origin.DistanceTo(point)
        except:
            continue
        if min_dist is None or d < min_dist:
            min_dist = d
            nearest = c
    return nearest


def pick_two_elements():
    """Get two connectable elements from current selection or by picking."""
    # try current selection first
    sel_ids = list(uidoc.Selection.GetElementIds())
    elements = []
    for elid in sel_ids:
        el = doc.GetElement(elid)
        if el and has_connectors(el):
            elements.append(el)

    # if not exactly two, ask user to pick
    if len(elements) < 2:
        filt = ConnectorSelectionFilter()
        ref1 = uidoc.Selection.PickObject(
            ObjectType.Element,
            filt,
            u"Выберите первый инженерный элемент"
        )
        ref2 = uidoc.Selection.PickObject(
            ObjectType.Element,
            filt,
            u"Выберите второй инженерный элемент"
        )
        e1 = doc.GetElement(ref1.ElementId)
        e2 = doc.GetElement(ref2.ElementId)
        return e1, e2

    if len(elements) > 2:
        elements = elements[:2]

    return elements[0], elements[1]


def main():
    # pick elements
    e1, e2 = pick_two_elements()

    # choose free connectors (если один свободный — выберется автоматически)
    conn1 = choose_connector(e1, u"Выберите коннектор первого элемента")
    conn2 = choose_connector(e2, u"Выберите коннектор второго элемента")

    # по умолчанию двигаем второй элемент
    moving_el = e2
    moving_conn = conn2
    fixed_el = e1
    fixed_conn = conn1

    move_vec = fixed_conn.Origin - moving_conn.Origin

    t = DB.Transaction(doc, "Connect with move")
    t.Start()
    try:
        # геометрическое перемещение выбранного элемента
        if move_vec.GetLength() > 1.0e-6:
            DB.ElementTransformUtils.MoveElement(doc, moving_el.Id, move_vec)

        # после перемещения обновим ссылки на элементы и коннекторы
        moving_el = doc.GetElement(moving_el.Id)
        fixed_el = doc.GetElement(fixed_el.Id)

        # находим ближайшие коннекторы (на случай, если после перемещения
        # Revit переставил их или мы выбрали не идеально совпадающие)
        conn_moving_final = find_closest_connector(moving_el, fixed_conn.Origin)
        conn_fixed_final = find_closest_connector(fixed_el, fixed_conn.Origin)

        if conn_moving_final is None or conn_fixed_final is None:
            raise Exception("Не удалось найти коннекторы после перемещения.")

        # попытка создать соединительную муфту
        try:
            doc.Create.NewUnionFitting(conn_moving_final, conn_fixed_final)
        except:
            # резервный вариант — прямое соединение коннекторов
            try:
                conn_moving_final.ConnectTo(conn_fixed_final)
            except:
                raise

        t.Commit()
    except Exception as ex:
        t.RollBack()
        forms.alert(
            u"Не удалось соединить элементы.\n\nПодробности:\n{0}".format(ex),
            ok=True
        )


if __name__ == "__main__":
    main()
