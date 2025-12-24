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


def get_free_connectors(element):
    """Return list of free (not connected) connectors for element."""
    free_connectors = []
    all_connectors = get_connectors(element)
    for c in all_connectors:
        try:
            if not c.IsConnected:
                free_connectors.append(c)
        except:
            # если по какой-то причине IsConnected недоступен,
            # считаем коннектор свободным
            free_connectors.append(c)
    return free_connectors


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


def auto_best_connector_pair(e1, e2):
    """Автоматически подобрать пару свободных коннекторов с минимальным расстоянием.

    e1 — первый элемент (остаётся на месте),
    e2 — второй элемент (будет перемещаться).
    """
    free1 = get_free_connectors(e1)
    free2 = get_free_connectors(e2)

    if not free1:
        forms.alert(
            u"У первого элемента (Id {0}) нет свободных коннекторов.".format(
                e1.Id.IntegerValue
            ),
            ok=True,
            exitscript=True
        )

    if not free2:
        forms.alert(
            u"У второго элемента (Id {0}) нет свободных коннекторов.".format(
                e2.Id.IntegerValue
            ),
            ok=True,
            exitscript=True
        )

    min_dist = None
    best_pair = (None, None)

    for c1 in free1:
        for c2 in free2:
            try:
                d = c1.Origin.DistanceTo(c2.Origin)
            except:
                continue
            if min_dist is None or d < min_dist:
                min_dist = d
                best_pair = (c1, c2)

    return best_pair


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


def are_collinear_mepcurves(el1, el2, tolerance=0.01):
    """Проверить, являются ли два MEPCurve соосными (на одной линии)."""
    try:
        if not isinstance(el1, DB.MEPCurve) or not isinstance(el2, DB.MEPCurve):
            return False
        
        curve1 = el1.Location.Curve
        curve2 = el2.Location.Curve
        
        # Получаем направления
        dir1 = (curve1.GetEndPoint(1) - curve1.GetEndPoint(0)).Normalize()
        dir2 = (curve2.GetEndPoint(1) - curve2.GetEndPoint(0)).Normalize()
        
        # Проверяем параллельность (dot product близок к 1 или -1)
        dot = abs(dir1.DotProduct(dir2))
        if dot < 0.999:
            return False
        
        # Проверяем, лежат ли на одной прямой
        # Вектор между точками должен быть параллелен направлению
        p1 = curve1.GetEndPoint(0)
        p2 = curve2.GetEndPoint(0)
        vec_between = p2 - p1
        
        if vec_between.GetLength() < 1e-6:
            return True
            
        vec_between_norm = vec_between.Normalize()
        dot2 = abs(vec_between_norm.DotProduct(dir1))
        
        return dot2 > 0.999
    except:
        return False


def merge_collinear_mepcurves(el1, el2, conn1, conn2):
    """Объединить два соосных MEPCurve в один, удлинив первый до конца второго."""
    try:
        curve1 = el1.Location.Curve
        curve2 = el2.Location.Curve
        
        # Точки первого элемента
        p1_start = curve1.GetEndPoint(0)
        p1_end = curve1.GetEndPoint(1)
        
        # Точки второго элемента
        p2_start = curve2.GetEndPoint(0)
        p2_end = curve2.GetEndPoint(1)
        
        # Определяем, какой конец первого элемента ближе к второму
        # conn1 - коннектор первого элемента, который мы соединяем
        conn1_origin = conn1.Origin
        
        # Определяем, какой конец второго элемента дальше от точки соединения
        # (это будет новый конец объединённой трубы)
        dist_p2_start = conn1_origin.DistanceTo(p2_start)
        dist_p2_end = conn1_origin.DistanceTo(p2_end)
        
        if dist_p2_start > dist_p2_end:
            new_far_point = p2_start
        else:
            new_far_point = p2_end
        
        # Определяем, какой конец первого элемента соединяется
        dist_p1_start = conn1_origin.DistanceTo(p1_start)
        dist_p1_end = conn1_origin.DistanceTo(p1_end)
        
        if dist_p1_start < dist_p1_end:
            # Соединяем через start первого элемента
            new_start = new_far_point
            new_end = p1_end
        else:
            # Соединяем через end первого элемента
            new_start = p1_start
            new_end = new_far_point
        
        # Создаём новую линию
        new_line = DB.Line.CreateBound(new_start, new_end)
        
        # Устанавливаем новую геометрию первому элементу
        el1.Location.Curve = new_line
        
        # Удаляем второй элемент
        doc.Delete(el2.Id)
        
        return True
    except Exception as ex:
        return False


def main():
    # pick elements
    e1, e2 = pick_two_elements()

    # автоматический подбор пары свободных коннекторов с минимальным расстоянием
    conn1, conn2 = auto_best_connector_pair(e1, e2)

    if conn1 is None or conn2 is None:
        forms.alert(
            u"Не удалось подобрать пару свободных коннекторов для соединения.",
            ok=True,
            exitscript=True
        )

    # по умолчанию двигаем второй элемент
    moving_el = e2
    moving_conn = conn2
    fixed_el = e1
    fixed_conn = conn1

    move_vec = fixed_conn.Origin - moving_conn.Origin

    t = DB.Transaction(doc, "Connect with move (auto pair)")
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

        # Проверяем, являются ли элементы соосными MEPCurve
        if are_collinear_mepcurves(fixed_el, moving_el):
            # Объединяем в один элемент
            if not merge_collinear_mepcurves(fixed_el, moving_el, conn_fixed_final, conn_moving_final):
                raise Exception("Не удалось объединить соосные элементы.")
        else:
            # Прямое соединение коннекторов
            try:
                conn_moving_final.ConnectTo(conn_fixed_final)
            except:
                # Если прямое соединение не сработало, пробуем муфту
                try:
                    doc.Create.NewUnionFitting(conn_moving_final, conn_fixed_final)
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
