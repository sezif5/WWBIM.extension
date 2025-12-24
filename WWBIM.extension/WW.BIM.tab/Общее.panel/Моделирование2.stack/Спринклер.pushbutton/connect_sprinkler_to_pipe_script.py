# -*- coding: utf-8 -*-
from __future__ import division

from pyrevit import revit, DB, UI, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = revit.doc
uidoc = revit.uidoc

# Константы
TAP_FAMILY_NAME = u"ADSK_СтальСварка_Врезка"
BRANCH_DIAMETER_MM = 15.0  # Диаметр отвода в мм


class PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            if isinstance(element, DB.Plumbing.Pipe):
                return True
            cat = element.Category
            if isinstance(element, DB.MEPCurve) and cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_PipeCurves):
                return True
        except Exception:
            pass
        return False

    def AllowReference(self, reference, point):
        return True


class SprinklerSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            cat = element.Category
            if cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_Sprinklers):
                return True
        except Exception:
            pass
        return False

    def AllowReference(self, reference, point):
        return True


def get_mep_connectors(element):
    """Вернуть список коннекторов элемента."""
    connectors = []
    cm = None

    mep_model = getattr(element, "MEPModel", None)
    if mep_model:
        try:
            cm = mep_model.ConnectorManager
        except Exception:
            cm = None

    if cm is None and hasattr(element, "ConnectorManager"):
        try:
            cm = element.ConnectorManager
        except Exception:
            cm = None

    if not cm:
        return connectors

    try:
        for c in cm.Connectors:
            connectors.append(c)
    except Exception:
        pass

    return connectors


def get_sprinkler_connector(sprinkler):
    """Получить трубопроводный коннектор спринклера."""
    connectors = get_mep_connectors(sprinkler)
    if not connectors:
        return None

    for c in connectors:
        try:
            if c.Domain == DB.Domain.DomainPiping:
                return c
        except Exception:
            pass

    return connectors[0] if connectors else None


def project_point_to_curve(curve, point):
    """Ортогональная проекция точки на кривую."""
    try:
        res = curve.Project(point)
        if res:
            return res.XYZPoint
    except Exception:
        pass

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    if point.DistanceTo(p0) < point.DistanceTo(p1):
        return p0
    return p1


def get_open_connector(element):
    """Найти незанятый (открытый) коннектор элемента."""
    connectors = get_mep_connectors(element)
    for c in connectors:
        try:
            if not c.IsConnected:
                return c
        except Exception:
            pass
    return None


def find_tap_family_symbol(family_name):
    """Найти типоразмер семейства врезки по имени."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_PipeFitting)
    
    for symbol in collector:
        try:
            family = symbol.Family
            if family and family.Name == family_name:
                return symbol
        except Exception:
            pass
    
    return None


def ensure_tap_routing_preference(pipe_type, tap_symbol):
    """Убедиться что в настройках маршрутизации есть врезка."""
    try:
        rpm = pipe_type.RoutingPreferenceManager
        if rpm is None:
            return False
        
        # Проверяем есть ли уже правило для Junction (отвод/врезка)
        junction_rule_count = rpm.GetNumberOfRules(DB.Plumbing.RoutingPreferenceRuleGroupType.Junctions)
        
        # Если нет правил для отводов, добавляем
        if junction_rule_count == 0:
            # Создаём правило для врезки
            rule = DB.Plumbing.RoutingPreferenceRule(tap_symbol.Id, u"Врезка")
            rpm.AddRule(DB.Plumbing.RoutingPreferenceRuleGroupType.Junctions, rule)
            return True
        
        return True
    except Exception:
        return False


def main():
    # 1. Выбор трубы
    pipe_ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        PipeSelectionFilter(),
        u"Выберите трубу"
    )
    pipe = doc.GetElement(pipe_ref.ElementId)

    # 2. Выбор спринклера
    sprinkler_ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        SprinklerSelectionFilter(),
        u"Выберите спринклер"
    )
    sprinkler = doc.GetElement(sprinkler_ref.ElementId)

    # 3. Получаем геометрию трубы
    loc = pipe.Location
    if not hasattr(loc, "Curve") or loc.Curve is None:
        forms.alert(u"Выбранный элемент не является линейной трубой.", exitscript=True)

    curve = loc.Curve

    # 4. Находим коннектор спринклера
    spr_conn = get_sprinkler_connector(sprinkler)
    if spr_conn is None:
        forms.alert(u"У выбранного спринклера не найден коннектор.", exitscript=True)

    if spr_conn.IsConnected:
        forms.alert(u"Спринклер уже подключён к системе. Сначала отсоедините его.", exitscript=True)

    spr_origin = spr_conn.Origin
    
    # 5. Проекция коннектора спринклера на ось трубы
    proj_point = project_point_to_curve(curve, spr_origin)
    
    # Проверяем расстояние от спринклера до трубы
    distance = spr_origin.DistanceTo(proj_point)
    if distance < 0.05:
        forms.alert(u"Спринклер слишком близко к трубе.", exitscript=True)

    # Проверяем расстояние от проекции до концов трубы
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dist_to_start = proj_point.DistanceTo(p0)
    dist_to_end = proj_point.DistanceTo(p1)
    
    if dist_to_start < 0.1 or dist_to_end < 0.1:
        forms.alert(u"Точка врезки слишком близко к концу трубы.", exitscript=True)

    # 6. Ищем семейство врезки
    tap_symbol = find_tap_family_symbol(TAP_FAMILY_NAME)
    if tap_symbol is None:
        forms.alert(
            u"Семейство врезки '{0}' не найдено в проекте.\n"
            u"Загрузите семейство и попробуйте снова.".format(TAP_FAMILY_NAME),
            exitscript=True
        )

    # Получаем параметры трубы
    pipe_type_id = pipe.GetTypeId()
    pipe_type = doc.GetElement(pipe_type_id)
    pipe_system_type_id = pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    pipe_level_id = pipe.get_Parameter(DB.BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()

    t = DB.Transaction(doc, u"Подключить спринклер к трубе")
    t.Start()
    try:
        # 7. Активируем символ врезки
        if not tap_symbol.IsActive:
            tap_symbol.Activate()
            doc.Regenerate()

        # 8. Добавляем врезку в настройки маршрутизации типа трубы (если нужно)
        ensure_tap_routing_preference(pipe_type, tap_symbol)
        doc.Regenerate()

        # 9. Создаём вертикальную трубу от точки на основной трубе
        temp_length = 200.0 / 304.8  # 200мм
        
        if spr_origin.Z < proj_point.Z:
            temp_end = DB.XYZ(proj_point.X, proj_point.Y, proj_point.Z - temp_length)
        else:
            temp_end = DB.XYZ(proj_point.X, proj_point.Y, proj_point.Z + temp_length)
        
        branch_pipe = DB.Plumbing.Pipe.Create(
            doc,
            pipe_system_type_id,
            pipe_type_id,
            pipe_level_id,
            proj_point,
            temp_end
        )
        
        # Устанавливаем диаметр 15мм
        branch_diameter_feet = BRANCH_DIAMETER_MM / 304.8
        try:
            branch_pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(branch_diameter_feet)
        except Exception:
            pass
        
        doc.Regenerate()
        
        # 10. Находим коннектор вертикальной трубы у основной
        branch_connectors = get_mep_connectors(branch_pipe)
        branch_conn_at_main = None
        branch_conn_at_end = None
        
        for c in branch_connectors:
            if c.Origin.DistanceTo(proj_point) < 0.01:
                branch_conn_at_main = c
            else:
                branch_conn_at_end = c
        
        # 11. Создаём врезку через NewTakeoffFitting
        tap_fitting = None
        error_msg = None
        
        if branch_conn_at_main:
            try:
                tap_fitting = doc.Create.NewTakeoffFitting(branch_conn_at_main, pipe)
            except Exception as ex:
                error_msg = str(ex)
        
        if tap_fitting is None:
            # Врезка не создалась - удаляем временную трубу и откатываем
            doc.Delete(branch_pipe.Id)
            t.RollBack()
            forms.alert(
                u"Врезка НЕ создана.\n\n"
                u"Ошибка: {0}\n\n"
                u"Создайте врезку вручную.".format(error_msg or u"неизвестно"),
                exitscript=True
            )
        
        # 12. Удаляем временную трубу
        doc.Delete(branch_pipe.Id)
        doc.Regenerate()
        
        # 13. Находим открытый коннектор врезки
        tap_open_conn = get_open_connector(tap_fitting)
        
        if tap_open_conn is None:
            t.Commit()
            forms.alert(
                u"Врезка создана, но нет открытого коннектора.\n"
                u"Подключите спринклер вручную.",
                ok=True
            )
            return
        
        # 14. Перемещаем спринклер к коннектору врезки
        tap_conn_origin = tap_open_conn.Origin
        move_vec = tap_conn_origin - spr_origin
        
        if move_vec.GetLength() > 0.001:
            DB.ElementTransformUtils.MoveElement(doc, sprinkler.Id, move_vec)
            sprinkler = doc.GetElement(sprinkler.Id)
            spr_conn = get_sprinkler_connector(sprinkler)
        
        # 15. Соединяем спринклер с врезкой
        try:
            tap_open_conn.ConnectTo(spr_conn)
        except Exception:
            pass
        
        t.Commit()

    except Exception as ex:
        t.RollBack()
        forms.alert(u"Ошибка: {0}".format(ex), ok=True)


if __name__ == '__main__':
    main()



