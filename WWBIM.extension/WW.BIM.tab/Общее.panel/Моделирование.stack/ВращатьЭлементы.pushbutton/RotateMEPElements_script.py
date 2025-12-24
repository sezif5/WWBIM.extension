
# -*- coding: utf-8 -*-
import sys
import math
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
clr.AddReference("System.Xml")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FamilyInstance, MEPCurve
from System.Collections.Generic import List
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader

# Получение документа
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


def get_connector_manager(element):
    """Вернуть ConnectorManager для MEP-элемента."""
    cm = None
    if isinstance(element, FamilyInstance):
        try:
            mep_model = element.MEPModel
            if mep_model:
                cm = mep_model.ConnectorManager
        except:
            cm = None
    if cm is None and isinstance(element, MEPCurve):
        try:
            cm = element.ConnectorManager
        except:
            cm = None
    return cm


def get_connectors(element):
    cm = get_connector_manager(element)
    if not cm:
        return []
    conns = []
    try:
        for c in cm.Connectors:
            conns.append(c)
    except:
        pass
    return conns


def find_nearest_connector(element, ref_point):
    """Найти коннектор элемента, ближайший к точке выбора."""
    connectors = get_connectors(element)
    if not connectors:
        return None
    nearest = None
    min_dist = None
    for c in connectors:
        try:
            d = (c.Origin - ref_point).GetLength()
        except:
            continue
        if min_dist is None or d < min_dist:
            min_dist = d
            nearest = c
    return nearest


def get_axis_direction(axis_element, axis_connector):
    """Получить направление оси вращения.

    Приоритет:
    1. Если осевой элемент - MEPCurve, берём направление её кривой.
    2. Иначе пытаемся найти MEPCurve среди подключенных к осевому коннектору.
    3. В качестве запаса используем базисы системы координат коннектора.
    """
    # 1. Направление самого осевого элемента
    try:
        if isinstance(axis_element, MEPCurve):
            curve = axis_element.Location.Curve
            vec = curve.GetEndPoint(1) - curve.GetEndPoint(0)
            if vec.GetLength() > 1e-6:
                return vec.Normalize()
    except:
        pass

    # 2. Ищем подключённый MEPCurve
    try:
        refs = list(axis_connector.AllRefs)
    except:
        refs = []

    for rc in refs:
        try:
            owner = rc.Owner
        except:
            owner = None
        if owner is None:
            continue
        try:
            if isinstance(owner, MEPCurve):
                curve = owner.Location.Curve
                vec = curve.GetEndPoint(1) - curve.GetEndPoint(0)
                if vec.GetLength() > 1e-6:
                    return vec.Normalize()
        except:
            continue

    # 3. Падаем обратно на локальную СК коннектора
    try:
        cs = axis_connector.CoordinateSystem
    except:
        cs = None
    if cs is not None:
        for v in [cs.BasisZ, cs.BasisX, cs.BasisY]:
            try:
                if v and v.GetLength() > 1e-6:
                    return v.Normalize()
            except:
                continue

    return XYZ.BasisZ


def is_connector_on_axis(connector, axis_dir):
    """Проверить, направлен ли коннектор вдоль оси вращения."""
    try:
        axis = axis_dir.Normalize()
    except:
        axis = axis_dir

    try:
        cs = connector.CoordinateSystem
        if cs is not None:
            v = cs.BasisZ
            if v is not None and v.GetLength() > 1e-6:
                dir_vec = v.Normalize()
                dot = abs(dir_vec.DotProduct(axis))
                if dot > 0.95:
                    return True
    except:
        pass

    return False


def is_connected_to_element(connector, element_id):
    """Проверить, подключён ли коннектор к указанному элементу."""
    try:
        refs = list(connector.AllRefs)
    except:
        return False

    for ref_conn in refs:
        try:
            owner = ref_conn.Owner
            if owner is not None and owner.Id == element_id:
                return True
        except:
            pass

    return False


def collect_elements_to_rotate(element_to_rotate, axis_element, axis_connector,
                               axis_dir, rotate_connected, include_longitudinal):
    """Собрать список элементов для вращения.

    rotate_connected:
        False  -> вращается только выбранный элемент.
        True   -> добавляем подключенные элементы (ветка сети).

    include_longitudinal:
        False  -> не добавляем элементы, идущие вдоль оси вращения (продольные).
        True   -> включаем также и продольные элементы.
    """
    # Только выбранный элемент
    if not rotate_connected:
        ids = List[ElementId]()
        ids.Add(element_to_rotate.Id)
        return ids

    try:
        axis_dir = axis_dir.Normalize()
    except:
        pass

    visited_ids = set()
    result_ids = List[ElementId]()

    visited_ids.add(element_to_rotate.Id.IntegerValue)
    result_ids.Add(element_to_rotate.Id)

    # Множество ID элементов в продольных ветках (не вращаем их)
    longitudinal_branch_ids = set()
    
    # Осевой элемент всегда в продольной ветке
    longitudinal_branch_ids.add(axis_element.Id.IntegerValue)

    # Определяем продольные коннекторы вращаемого элемента
    # Коннектор продольный если:
    # 1. Он направлен вдоль оси вращения (для тройников, труб)
    # 2. ИЛИ он подключён к осевому элементу (для отводов)
    if not include_longitudinal:
        cm = get_connector_manager(element_to_rotate)
        if cm:
            try:
                connectors = list(cm.Connectors)
            except:
                connectors = []

            for conn in connectors:
                is_longitudinal = False
                
                # Критерий 1: коннектор направлен вдоль оси
                if is_connector_on_axis(conn, axis_dir):
                    is_longitudinal = True
                
                # Критерий 2: коннектор подключён к осевому элементу
                if is_connected_to_element(conn, axis_element.Id):
                    is_longitudinal = True
                
                if is_longitudinal:
                    # Помечаем все элементы через этот коннектор как продольную ветку
                    try:
                        refs = list(conn.AllRefs)
                    except:
                        refs = []

                    for ref_conn in refs:
                        try:
                            owner = ref_conn.Owner
                        except:
                            owner = None
                        if owner is None:
                            continue
                        # Помечаем элемент и всю ветку за ним
                        mark_longitudinal_branch(owner, element_to_rotate.Id.IntegerValue, visited_ids.copy(), longitudinal_branch_ids)

    # Теперь собираем элементы для вращения, исключая продольные ветки
    to_visit = [element_to_rotate]

    while to_visit:
        current = to_visit.pop()
        cm = get_connector_manager(current)
        if not cm:
            continue

        try:
            connectors = list(cm.Connectors)
        except:
            connectors = []

        for conn in connectors:
            try:
                refs = list(conn.AllRefs)
            except:
                refs = []

            for ref_conn in refs:
                try:
                    owner = ref_conn.Owner
                except:
                    owner = None
                if owner is None:
                    continue

                # Не вращаем осевой элемент
                if owner.Id == axis_element.Id:
                    continue

                iid = owner.Id.IntegerValue
                if iid in visited_ids:
                    continue

                visited_ids.add(iid)

                # Если элемент в продольной ветке - пропускаем
                if iid in longitudinal_branch_ids:
                    continue

                # Добавляем элемент в список для вращения
                result_ids.Add(owner.Id)
                to_visit.append(owner)

    return result_ids


def mark_longitudinal_branch(start_element, exclude_element_id, visited, longitudinal_set):
    """Рекурсивно помечает все элементы в продольной ветке.
    
    exclude_element_id - ID элемента, который не нужно обходить (вращаемый элемент)
    """
    to_visit = [start_element]

    while to_visit:
        current = to_visit.pop()
        iid = current.Id.IntegerValue

        if iid in visited:
            continue
        visited.add(iid)
        longitudinal_set.add(iid)

        cm = get_connector_manager(current)
        if not cm:
            continue

        try:
            connectors = list(cm.Connectors)
        except:
            connectors = []

        for conn in connectors:
            try:
                refs = list(conn.AllRefs)
            except:
                refs = []

            for ref_conn in refs:
                try:
                    owner = ref_conn.Owner
                except:
                    owner = None
                if owner is None:
                    continue

                # Не возвращаемся к вращаемому элементу
                if owner.Id.IntegerValue == exclude_element_id:
                    continue

                if owner.Id.IntegerValue not in visited:
                    to_visit.append(owner)


def pick_elements():
    """Выбор элементов пользователем.

    Поддерживается предварительный выбор элемента для вращения.
    """
    sel = uidoc.Selection
    element_to_rotate = None

    # Предварительно выбранный элемент
    preselected_ids = list(sel.GetElementIds())
    if preselected_ids:
        element_to_rotate = doc.GetElement(preselected_ids[0])

    # Явный выбор элемента для вращения
    if element_to_rotate is None:
        try:
            ref1 = sel.PickObject(Selection.ObjectType.Element,
                                  "Выберите элемент для вращения")
        except:
            return None, None, None
        element_to_rotate = doc.GetElement(ref1.ElementId)

    # Элемент возле нужного коннектора
    try:
        ref2 = sel.PickObject(Selection.ObjectType.Element,
                              "Выберите ближайший элемент к коннектору, вокруг которого вы хотите вращать")
    except:
        return None, None, None

    axis_element = doc.GetElement(ref2.ElementId)
    axis_pick_point = ref2.GlobalPoint

    return element_to_rotate, axis_element, axis_pick_point


def show_settings_dialog(element_to_rotate, axis_element, axis_connector, axis_dir):
    """Показать окно настроек. Вращение выполняется сразу по нажатию на кнопки направления."""
    xaml = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ВРАЩЕНИЕ МЕР ЭЛЕМЕНТОВ"
        SizeToContent="WidthAndHeight"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Topmost="True"
        ShowInTaskbar="False">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Первая строка: направление и угол -->
        <StackPanel Orientation="Horizontal" Grid.Row="0" VerticalAlignment="Center">
            <Button x:Name="ClockwiseButton"
                    Content="По часовой"
                    Width="100"
                    Margin="0,0,5,0"/>
            <TextBox x:Name="AngleBox"
                     Width="50"
                     Text="90"
                     Margin="0,0,5,0"
                     HorizontalContentAlignment="Center"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,5,0">
                <Button x:Name="PlusButton"
                        Content="+"
                        Width="24"
                        Margin="0,0,2,0"/>
                <Button x:Name="MinusButton"
                        Content="-"
                        Width="24"/>
            </StackPanel>
            <Button x:Name="ResetButton"
                    Content="⟳"
                    Width="28"
                    Margin="0,0,5,0"/>
            <Button x:Name="CounterClockwiseButton"
                    Content="Против часовой"
                    Width="120"/>
        </StackPanel>

        <!-- Вторая строка: чекбоксы -->
        <StackPanel Grid.Row="1" Margin="0,10,0,0">
            <CheckBox x:Name="RotateConnectedCheck"
                      Content="Вращать подключенные элементы"
                      IsChecked="True"
                      Margin="0,0,0,5"/>
            <CheckBox x:Name="RotateLongitudinalCheck"
                      Content="Включая продольно подключенные элементы"/>
        </StackPanel>

        <!-- Третья строка: кнопки управления -->
        <StackPanel Grid.Row="2"
                    Orientation="Horizontal"
                    HorizontalAlignment="Right"
                    Margin="0,10,0,0">
            <Button x:Name="OkButton"
                    Content="Принять"
                    Width="75"
                    Margin="0,0,5,0"
                    IsDefault="True"/>
            <Button x:Name="CancelButton"
                    Content="Отмена"
                    Width="75"
                    IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"""

    reader = XmlReader.Create(StringReader(xaml))
    window = XamlReader.Load(reader)

    clockwise_btn = window.FindName("ClockwiseButton")
    counter_btn = window.FindName("CounterClockwiseButton")
    angle_box = window.FindName("AngleBox")
    plus_btn = window.FindName("PlusButton")
    minus_btn = window.FindName("MinusButton")
    reset_btn = window.FindName("ResetButton")
    rotate_connected_chk = window.FindName("RotateConnectedCheck")
    rotate_longitudinal_chk = window.FindName("RotateLongitudinalCheck")
    ok_button = window.FindName("OkButton")
    cancel_button = window.FindName("CancelButton")

    try:
        axis_line = Line.CreateUnbound(axis_connector.Origin, axis_dir)
    except:
        axis_line = Line.CreateUnbound(axis_connector.Origin, XYZ.BasisZ)

    # Группа транзакций для возможности отката всех изменений
    tg = TransactionGroup(doc, "Вращение MEP элементов")
    tg.Start()

    def parse_angle():
        try:
            val = float(angle_box.Text.replace(',', '.'))
        except:
            val = 0.0
        return val

    def rotate_step(clockwise):
        angle_deg = parse_angle()
        if abs(angle_deg) < 1e-6:
            return

        angle_rad = math.radians(angle_deg)
        if clockwise:
            angle_rad = -angle_rad

        rotate_connected = bool(rotate_connected_chk.IsChecked)
        include_longitudinal = bool(rotate_longitudinal_chk.IsChecked)

        ids_to_rotate = collect_elements_to_rotate(
            element_to_rotate,
            axis_element,
            axis_connector,
            axis_dir,
            rotate_connected,
            include_longitudinal
        )

        if ids_to_rotate is None or ids_to_rotate.Count == 0:
            try:
                TaskDialog.Show("Вращение", "Не удалось определить элементы для вращения.")
            except:
                pass
            return

        t = Transaction(doc, "Вращение MEP элементов")
        try:
            t.Start()
            ElementTransformUtils.RotateElements(doc, ids_to_rotate, axis_line, angle_rad)
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            try:
                TaskDialog.Show("Вращение", "Ошибка при выполнении вращения: {0}".format(str(ex)))
            except:
                pass

    # Обработчики направления
    def on_clockwise(sender, args):
        rotate_step(True)

    def on_counter(sender, args):
        rotate_step(False)

    clockwise_btn.Click += on_clockwise
    counter_btn.Click += on_counter

    # Изменение угла
    def change_angle(delta):
        val = parse_angle()
        val += delta
        if val < 0:
            val = 0
        angle_box.Text = str(int(val))

    def on_plus(sender, args):
        change_angle(15.0)

    def on_minus(sender, args):
        change_angle(-15.0)

    def on_reset(sender, args):
        angle_box.Text = "90"

    plus_btn.Click += on_plus
    minus_btn.Click += on_minus
    reset_btn.Click += on_reset

    # Кнопки управления
    def on_ok(sender, args):
        # Применить все изменения
        if tg.HasStarted():
            tg.Assimilate()
        window.DialogResult = True
        window.Close()

    def on_cancel(sender, args):
        # Откатить все изменения
        if tg.HasStarted():
            tg.RollBack()
        window.DialogResult = False
        window.Close()

    ok_button.Click += on_ok
    cancel_button.Click += on_cancel

    result = window.ShowDialog()
    
    # Если окно закрыто не через кнопки (например, крестиком), откатываем
    if tg.HasStarted():
        tg.RollBack()


def main():
    element_to_rotate, axis_element, axis_pick_point = pick_elements()
    if element_to_rotate is None or axis_element is None:
        return

    axis_connector = find_nearest_connector(axis_element, axis_pick_point)
    if axis_connector is None:
        TaskDialog.Show("Вращение", "Не удалось найти коннектор у выбранного элемента.")
        return

    axis_dir = get_axis_direction(axis_element, axis_connector)

    show_settings_dialog(element_to_rotate, axis_element, axis_connector, axis_dir)


if __name__ == "__main__":
    main()
