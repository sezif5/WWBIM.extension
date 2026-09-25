# -*- coding: utf-8 -*-
from __future__ import print_function, division

import io
import math
import traceback
from collections import deque
from datetime import datetime

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

FT_TO_M = 0.3048
M_TO_FT = 1.0 / FT_TO_M
FT2_TO_M2 = FT_TO_M * FT_TO_M
FT3_TO_M3 = FT_TO_M * FT_TO_M * FT_TO_M
DEFAULT_CELL_M = 0.25
MIN_CELL_M = 0.10
MAX_CELL_M = 1.00
MAX_GRID_CELLS = 1000000
GRID_MARGIN_M = 2.0
MIN_BAND_M = 0.05
EDGE_MERGE_M = 0.03
MIN_COMPONENT_M2 = 2.0
MIN_FOUNDATION_PLATE_M2 = 10.0
PLATE_FALLBACK_GROUP_M = 0.20
PLATE_LEVEL_OFFSET_GROUP_M = 0.50
DS_NAME_PREFIX = u"РАСЧЁТ_ОбъёмЗдания"
MODE_HYBRID = u"hybrid"
MODE_PLATES = u"plates"

TEXT_TYPE = type(u"")


class WallRecord(object):
    def __init__(self, elem, points, width, bounds):
        self.elem = elem
        self.points = points
        self.width = width
        self.xmin, self.ymin, self.xmax, self.ymax, self.zmin, self.zmax = bounds


class FloorRecord(object):
    def __init__(self, elem, elevation, polygon, level_key):
        self.elem = elem
        self.elevation = elevation
        self.polygon = polygon
        self.level_key = level_key


class Grid(object):
    def __init__(self, xmin, ymin, xmax, ymax, requested_cell):
        self.requested_cell = requested_cell
        self.cell = requested_cell
        width = max(requested_cell, xmax - xmin)
        height = max(requested_cell, ymax - ymin)
        estimated = int(math.ceil(width / self.cell)) * int(math.ceil(height / self.cell))
        if estimated > MAX_GRID_CELLS:
            self.cell *= math.sqrt(estimated / float(MAX_GRID_CELLS))
            step = 0.05 * M_TO_FT
            self.cell = math.ceil(self.cell / step) * step
        self.xmin = math.floor(xmin / self.cell) * self.cell
        self.ymin = math.floor(ymin / self.cell) * self.cell
        self.nx = max(1, int(math.ceil((xmax - self.xmin) / self.cell)))
        self.ny = max(1, int(math.ceil((ymax - self.ymin) / self.cell)))

    def center(self, i, j):
        return self.xmin + (i + 0.5) * self.cell, self.ymin + (j + 0.5) * self.cell

    def index_bounds(self, xmin, ymin, xmax, ymax):
        i0 = max(0, int(math.floor((xmin - self.xmin) / self.cell)))
        j0 = max(0, int(math.floor((ymin - self.ymin) / self.cell)))
        i1 = min(self.nx - 1, int(math.floor((xmax - self.xmin) / self.cell)))
        j1 = min(self.ny - 1, int(math.floor((ymax - self.ymin) / self.cell)))
        return i0, j0, i1, j1


class LevelChoice(object):
    def __init__(self, level):
        self.level = level
        try:
            elevation = level.Elevation * FT_TO_M
        except Exception:
            elevation = 0.0
        self.label = u"{} ({:+.2f} м)".format(to_unicode(level.Name), elevation)


class NoGroundChoice(object):
    def __init__(self):
        self.level = None
        self.label = u"— нет (всё считать надземным) —"


class ModeChoice(object):
    def __init__(self, key, label, description):
        self.key = key
        self.label = label
        self.description = description

    @property
    def display_name(self):
        return u"{} — {}".format(self.label, self.description)


def to_unicode(value):
    if value is None:
        return u""
    try:
        return TEXT_TYPE(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def category_id(elem):
    try:
        return elem.Category.Id.IntegerValue
    except Exception:
        return None


def element_id(elem):
    try:
        return elem.Id.IntegerValue
    except Exception:
        return -1


def element_bounds(elem):
    try:
        bbox = elem.get_BoundingBox(None)
        if bbox is None:
            return None
        return bbox.Min.X, bbox.Min.Y, bbox.Max.X, bbox.Max.Y, bbox.Min.Z, bbox.Max.Z
    except Exception:
        return None


def is_exterior_wall(wall):
    try:
        param = wall.WallType.get_Parameter(DB.BuiltInParameter.FUNCTION_PARAM)
        return param is not None and param.AsInteger() == int(DB.WallFunction.Exterior)
    except Exception:
        return False


def geometry_options():
    try:
        options = DB.Options()
        options.DetailLevel = DB.ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = False
        options.ComputeReferences = False
        return options
    except Exception:
        return None


def collect_solids_recursive(geometry, result):
    if geometry is None:
        return
    try:
        iterator = geometry.GetEnumerator()
        while iterator.MoveNext():
            item = iterator.Current
            try:
                if isinstance(item, DB.Solid):
                    if item.Faces.Size > 0 and item.Volume > 1e-8:
                        result.append(item)
                elif isinstance(item, DB.GeometryInstance):
                    collect_solids_recursive(item.GetInstanceGeometry(), result)
                elif isinstance(item, DB.GeometryElement):
                    collect_solids_recursive(item, result)
            except Exception:
                pass
    except Exception:
        pass


def element_solids(elem, options):
    solids = []
    try:
        collect_solids_recursive(elem.get_Geometry(options), solids)
    except Exception:
        pass
    return solids


def curve_points(curve):
    points = []
    try:
        points = list(curve.Tessellate())
    except Exception:
        pass
    if len(points) < 2:
        try:
            points = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
        except Exception:
            points = []
    return points


def loop_points(loop):
    points = []
    try:
        for curve in loop:
            tessellated = list(curve.Tessellate())
            if not tessellated:
                continue
            if points:
                tessellated = tessellated[1:]
            points.extend(tessellated)
    except Exception:
        return []
    if len(points) > 2:
        try:
            if points[0].DistanceTo(points[-1]) < 1e-7:
                points.pop()
        except Exception:
            pass
    return [(p.X, p.Y) for p in points]


def polygon_area(polygon):
    if len(polygon) < 3:
        return 0.0
    area = 0.0
    for index in range(len(polygon)):
        x1, y1 = polygon[index]
        x2, y2 = polygon[(index + 1) % len(polygon)]
        area += x1 * y2 - x2 * y1
    return 0.5 * area


def build_wall_records(elements):
    records = []
    all_walls = []
    for elem in elements:
        try:
            if isinstance(elem, DB.Wall):
                all_walls.append(elem)
        except Exception:
            pass
    exterior = [wall for wall in all_walls if is_exterior_wall(wall)]
    walls = exterior if len(exterior) >= 3 else all_walls
    for wall in walls:
        bounds = element_bounds(wall)
        if bounds is None:
            continue
        try:
            curve = wall.Location.Curve
        except Exception:
            curve = None
        if curve is None:
            continue
        points = curve_points(curve)
        if len(points) < 2:
            continue
        try:
            width = float(wall.Width)
        except Exception:
            width = 0.20 * M_TO_FT
        if width < 0.05 * M_TO_FT:
            width = 0.20 * M_TO_FT
        records.append(WallRecord(wall, points, width, bounds))
    return records, len(exterior) >= 3


def plate_level_key(elem):
    try:
        level_id = elem.LevelId
        if level_id is not None and level_id != DB.ElementId.InvalidElementId:
            return level_id.IntegerValue
    except Exception:
        pass
    parameter_names = (
        "LEVEL_PARAM",
        "ROOF_BASE_LEVEL_PARAM",
        "FAMILY_LEVEL_PARAM",
    )
    for parameter_name in parameter_names:
        try:
            built_in = getattr(DB.BuiltInParameter, parameter_name)
            parameter = elem.get_Parameter(built_in)
            if parameter is None:
                continue
            level_id = parameter.AsElementId()
            if level_id is not None and level_id != DB.ElementId.InvalidElementId:
                return level_id.IntegerValue
        except Exception:
            pass
    return None


def build_floor_records(elements, options):
    records = []
    plate_categories = {
        int(DB.BuiltInCategory.OST_Floors),
        int(DB.BuiltInCategory.OST_Roofs),
        int(DB.BuiltInCategory.OST_StructuralFoundation),
    }
    for elem in elements:
        if category_id(elem) not in plate_categories:
            continue
        bounds = element_bounds(elem)
        if bounds is None:
            continue
        if category_id(elem) == int(DB.BuiltInCategory.OST_StructuralFoundation):
            footprint_m2 = (bounds[2] - bounds[0]) * (bounds[3] - bounds[1]) * FT2_TO_M2
            if footprint_m2 < MIN_FOUNDATION_PLATE_M2:
                continue
        element_top = bounds[5]
        level_key = plate_level_key(elem)
        for solid in element_solids(elem, options):
            faces = []
            try:
                for face in solid.Faces:
                    try:
                        if isinstance(face, DB.PlanarFace) and face.FaceNormal.Z > 0.85:
                            faces.append((face.Origin.Z, face))
                    except Exception:
                        pass
            except Exception:
                pass
            if not faces:
                continue
            highest = max(item[0] for item in faces)
            for elevation, face in faces:
                if highest - elevation > 0.01 * M_TO_FT:
                    continue
                if element_top - elevation > 0.20 * M_TO_FT:
                    continue
                loops = []
                try:
                    loops = list(face.GetEdgesAsCurveLoops())
                except Exception:
                    pass
                measured = []
                for loop in loops:
                    polygon = loop_points(loop)
                    area = abs(polygon_area(polygon))
                    if area > 0.25 * M_TO_FT * M_TO_FT:
                        measured.append((area, polygon))
                if measured:
                    measured.sort(key=lambda item: -item[0])
                    records.append(FloorRecord(elem, element_top, measured[0][1], level_key))
    return records


def build_roof_triangles(elements, options):
    triangles = []
    roof_cat = int(DB.BuiltInCategory.OST_Roofs)
    for elem in elements:
        if category_id(elem) != roof_cat:
            continue
        for solid in element_solids(elem, options):
            try:
                faces = list(solid.Faces)
            except Exception:
                faces = []
            for face in faces:
                try:
                    mesh = face.Triangulate()
                except Exception:
                    continue
                try:
                    count = mesh.NumTriangles
                except Exception:
                    count = 0
                for index in range(count):
                    try:
                        triangle = mesh.get_Triangle(index)
                        a = triangle.get_Vertex(0)
                        b = triangle.get_Vertex(1)
                        c = triangle.get_Vertex(2)
                        normal = (b - a).CrossProduct(c - a)
                        projected = abs((b.X - a.X) * (c.Y - a.Y) - (c.X - a.X) * (b.Y - a.Y))
                        if normal.Z > 1e-8 and projected > 1e-8:
                            triangles.append((a, b, c))
                    except Exception:
                        pass
    return triangles


def get_relevant_elements():
    selected = []
    try:
        selected = [elem for elem in revit.get_selection() if elem is not None]
    except Exception:
        pass
    relevant_ids = {
        int(DB.BuiltInCategory.OST_Walls),
        int(DB.BuiltInCategory.OST_Floors),
        int(DB.BuiltInCategory.OST_Roofs),
        int(DB.BuiltInCategory.OST_StructuralFoundation),
    }
    selected = [elem for elem in selected if category_id(elem) in relevant_ids]
    if selected:
        return selected, u"Текущее выделение"

    elements = []
    try:
        walls = list(DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType())
    except Exception:
        walls = []
    exterior = [wall for wall in walls if is_exterior_wall(wall)]
    elements.extend(exterior if len(exterior) >= 3 else walls)
    for bic in (
        DB.BuiltInCategory.OST_Floors,
        DB.BuiltInCategory.OST_Roofs,
        DB.BuiltInCategory.OST_StructuralFoundation,
    ):
        try:
            elements.extend(list(DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()))
        except Exception:
            pass
    return elements, u"Автоматический сбор"


def get_levels():
    try:
        levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level).WhereElementIsNotElementType())
    except Exception:
        levels = []
    levels.sort(key=lambda level: safe_level_elevation(level))
    return levels


def safe_level_elevation(level):
    try:
        return float(level.Elevation)
    except Exception:
        return 0.0


def ask_ground_level(levels):
    choices = list()
    choices.append(NoGroundChoice())
    for level in levels:
        choices.append(LevelChoice(level))
    picked = forms.SelectFromList.show(
        choices,
        name_attr="label",
        multiselect=False,
        title=u"Уровень земли (граница надземный/подземный)"
    )
    if picked is None:
        return None, False
    return picked.level, True


def ask_calculation_mode():
    choices = [
        ModeChoice(
            MODE_HYBRID,
            u"Гибридный",
            u"наружные стены с дополнением перекрытиями/крышами/фундаментными плитами"
        ),
        ModeChoice(
            MODE_PLATES,
            u"Упрощённый по плитам",
            u"только суммарные контуры перекрытий/крыш/фундаментных плит"
        ),
    ]
    picked = forms.SelectFromList.show(
        choices,
        name_attr="display_name",
        multiselect=False,
        title=u"Режим расчёта строительного объёма"
    )
    return picked


def ask_cell_size():
    text = forms.ask_for_string(
        default=u"0.25",
        prompt=u"Размер ячейки внешнего контура, м (0.25 — точно, 0.50 — быстрее)",
        title=u"Объём здания"
    )
    if text is None:
        return None
    try:
        value = float(to_unicode(text).replace(u",", u".").strip())
    except Exception:
        value = DEFAULT_CELL_M
    value = max(MIN_CELL_M, min(MAX_CELL_M, value))
    return value


def distance_sq_to_segment(px, py, ax, ay, bx, by):
    dx = bx - ax
    dy = by - ay
    length_sq = dx * dx + dy * dy
    if length_sq < 1e-16:
        return (px - ax) * (px - ax) + (py - ay) * (py - ay)
    parameter = ((px - ax) * dx + (py - ay) * dy) / length_sq
    parameter = max(0.0, min(1.0, parameter))
    qx = ax + parameter * dx
    qy = ay + parameter * dy
    return (px - qx) * (px - qx) + (py - qy) * (py - qy)


def rasterize_walls(records, grid, elevation):
    cells = set()
    active = 0
    vertical_tolerance = max(0.20 * M_TO_FT, 0.5 * grid.cell)
    for record in records:
        if elevation < record.zmin - vertical_tolerance or elevation > record.zmax + vertical_tolerance:
            continue
        active += 1
        radius = 0.5 * record.width + 0.55 * grid.cell
        radius_sq = radius * radius
        for index in range(len(record.points) - 1):
            a = record.points[index]
            b = record.points[index + 1]
            i0, j0, i1, j1 = grid.index_bounds(
                min(a.X, b.X) - radius,
                min(a.Y, b.Y) - radius,
                max(a.X, b.X) + radius,
                max(a.Y, b.Y) + radius,
            )
            for j in range(j0, j1 + 1):
                for i in range(i0, i1 + 1):
                    x, y = grid.center(i, j)
                    if distance_sq_to_segment(x, y, a.X, a.Y, b.X, b.Y) <= radius_sq:
                        cells.add((i, j))
    return cells, active


def connected_components(cells):
    remaining = set(cells)
    result = []
    while remaining:
        start = remaining.pop()
        component = {start}
        queue = deque([start])
        while queue:
            i, j = queue.popleft()
            for neighbor in ((i - 1, j), (i + 1, j), (i, j - 1), (i, j + 1)):
                if neighbor in remaining:
                    remaining.remove(neighbor)
                    component.add(neighbor)
                    queue.append(neighbor)
        result.append(component)
    return result


def dilate_cells(cells, grid):
    if not cells:
        return set()
    dilated = set(cells)
    for i, j in cells:
        for di in (-1, 0, 1):
            for dj in (-1, 0, 1):
                ni, nj = i + di, j + dj
                if 0 <= ni < grid.nx and 0 <= nj < grid.ny:
                    dilated.add((ni, nj))
    return dilated


def morphological_close(cells, grid):
    if not cells:
        return set()
    dilated = dilate_cells(cells, grid)
    eroded = set()
    for i, j in dilated:
        keep = True
        for di in (-1, 0, 1):
            for dj in (-1, 0, 1):
                if (i + di, j + dj) not in dilated:
                    keep = False
                    break
            if not keep:
                break
        if keep:
            eroded.add((i, j))
    return eroded if eroded else set(cells)


def fill_enclosed_cells(occupied, grid):
    if not occupied:
        return set()
    outside = set()
    queue = deque()
    for i in range(grid.nx):
        for cell in ((i, 0), (i, grid.ny - 1)):
            if cell not in occupied and cell not in outside:
                outside.add(cell)
                queue.append(cell)
    for j in range(grid.ny):
        for cell in ((0, j), (grid.nx - 1, j)):
            if cell not in occupied and cell not in outside:
                outside.add(cell)
                queue.append(cell)
    while queue:
        i, j = queue.popleft()
        for ni, nj in ((i - 1, j), (i + 1, j), (i, j - 1), (i, j + 1)):
            if ni < 0 or nj < 0 or ni >= grid.nx or nj >= grid.ny:
                continue
            cell = (ni, nj)
            if cell in occupied or cell in outside:
                continue
            outside.add(cell)
            queue.append(cell)
    filled = set()
    for j in range(grid.ny):
        for i in range(grid.nx):
            if (i, j) not in outside:
                filled.add((i, j))
    return filled


def enclosed_footprint(wall_cells, grid):
    if not wall_cells:
        return set(), 0
    wall_cells = morphological_close(wall_cells, grid)
    outside = set()
    queue = deque()
    for i in range(grid.nx):
        for cell in ((i, 0), (i, grid.ny - 1)):
            if cell not in wall_cells and cell not in outside:
                outside.add(cell)
                queue.append(cell)
    for j in range(grid.ny):
        for cell in ((0, j), (grid.nx - 1, j)):
            if cell not in wall_cells and cell not in outside:
                outside.add(cell)
                queue.append(cell)
    while queue:
        i, j = queue.popleft()
        for ni, nj in ((i - 1, j), (i + 1, j), (i, j - 1), (i, j + 1)):
            if ni < 0 or nj < 0 or ni >= grid.nx or nj >= grid.ny:
                continue
            cell = (ni, nj)
            if cell in wall_cells or cell in outside:
                continue
            outside.add(cell)
            queue.append(cell)
    inside = set()
    for j in range(grid.ny):
        for i in range(grid.nx):
            if (i, j) not in outside:
                inside.add((i, j))
    min_enclosed = max(4, int(math.ceil(MIN_COMPONENT_M2 / (grid.cell * grid.cell * FT2_TO_M2))))
    kept = set()
    enclosed_total = 0
    for component in connected_components(inside):
        enclosed = sum(1 for cell in component if cell not in wall_cells)
        if enclosed >= min_enclosed:
            kept.update(component)
            enclosed_total += enclosed
    return kept, enclosed_total


def point_in_polygon(x, y, polygon):
    inside = False
    count = len(polygon)
    if count < 3:
        return False
    j = count - 1
    for i in range(count):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        if ((yi > y) != (yj > y)):
            cross_x = (xj - xi) * (y - yi) / (yj - yi) + xi
            if x < cross_x:
                inside = not inside
        j = i
    return inside


def rasterize_polygon(polygon, grid, cells):
    if len(polygon) < 3:
        return
    xs = [point[0] for point in polygon]
    ys = [point[1] for point in polygon]
    i0, j0, i1, j1 = grid.index_bounds(min(xs), min(ys), max(xs), max(ys))
    for j in range(j0, j1 + 1):
        for i in range(i0, i1 + 1):
            x, y = grid.center(i, j)
            if point_in_polygon(x, y, polygon):
                cells.add((i, j))


def build_plate_cell_groups(records, grid, strict=False):
    if not records:
        return []
    keyed = {}
    unkeyed = []
    for record in records:
        if record.level_key is None:
            unkeyed.append(record)
        else:
            keyed.setdefault(record.level_key, []).append(record)

    groups = []
    level_offset_tolerance = PLATE_LEVEL_OFFSET_GROUP_M * M_TO_FT
    for level_records in keyed.values():
        level_groups = []
        for record in sorted(level_records, key=lambda item: item.elevation):
            if not level_groups or record.elevation - level_groups[-1][0] > level_offset_tolerance:
                level_groups.append([record.elevation, [record]])
            else:
                level_groups[-1][1].append(record)
                level_groups[-1][0] = max(level_groups[-1][0], record.elevation)
        groups.extend(level_groups)

    group_tolerance = PLATE_FALLBACK_GROUP_M * M_TO_FT
    for record in sorted(unkeyed, key=lambda item: item.elevation):
        matching = None
        for group in groups:
            if abs(record.elevation - group[0]) <= group_tolerance:
                matching = group
                break
        if matching is None:
            groups.append([record.elevation, [record]])
        else:
            matching[1].append(record)
            matching[0] = max(matching[0], record.elevation)
    groups.sort(key=lambda item: item[0])

    result = []
    for group_elevation, group_records in groups:
        cells = set()
        for record in group_records:
            rasterize_polygon(record.polygon, grid, cells)
        if not cells:
            continue
        if strict:
            filtered = cells
        else:
            components = connected_components(cells)
            largest_component = max(len(component) for component in components)
            component_minimum = max(4, int(largest_component * 0.02))
            filtered = set()
            for component in components:
                if len(component) >= component_minimum:
                    filtered.update(component)
        if filtered:
            contour = filtered if strict else morphological_close(filtered, grid)
            result.append((group_elevation, fill_enclosed_cells(contour, grid)))
    return result


def nearest_floor_cells(records, grid, elevation):
    candidates = build_plate_cell_groups(records, grid, strict=False)
    if not candidates:
        return set(), None

    largest_floor = max(len(item[1]) for item in candidates)
    substantial_minimum = max(4, int(largest_floor * 0.02))
    substantial = [item for item in candidates if len(item[1]) >= substantial_minimum]
    pool = substantial if substantial else candidates
    nearest = min(pool, key=lambda item: abs(item[0] - elevation))
    return nearest[1], nearest[0]


def choose_footprint(walls, floors, grid, elevation):
    wall_cells, active_walls = rasterize_walls(walls, grid, elevation)
    wall_footprint, enclosed_count = enclosed_footprint(wall_cells, grid)
    floor_cells, floor_elevation = nearest_floor_cells(floors, grid, elevation)
    source = u"Наружные стены"
    warnings = []

    wall_area = len(wall_footprint) * grid.cell * grid.cell
    floor_area = len(floor_cells) * grid.cell * grid.cell
    wall_valid = enclosed_count > 0 and wall_area > MIN_COMPONENT_M2 / FT2_TO_M2
    if wall_valid and floor_cells:
        ratio = wall_area / floor_area if floor_area > 1e-9 else 1.0
        if ratio < 0.55 or ratio > 1.80:
            hybrid = fill_enclosed_cells(morphological_close(wall_footprint | floor_cells | wall_cells, grid), grid)
            warnings.append(u"Контур стен существенно дополнен ближайшим перекрытием")
            return hybrid, u"Гибрид стен и перекрытия", active_walls, warnings
    if wall_valid:
        return wall_footprint, source, active_walls, warnings
    if floor_cells:
        hybrid = fill_enclosed_cells(morphological_close(floor_cells | wall_cells, grid), grid)
        warnings.append(u"Разрывы стен замкнуты ближайшим перекрытием")
        return hybrid, u"Гибрид стен и перекрытия", active_walls, warnings
    warnings.append(u"Не найден замкнутый контур стен или плиты")
    return set(), u"Нет контура", active_walls, warnings


def choose_plate_footprint(plates, grid, elevation):
    cells, plate_elevation = nearest_floor_cells(plates, grid, elevation)
    if not cells:
        return set(), u"Нет плиты", 0, [u"На отметке не найден контур перекрытия/крыши/фундаментной плиты"]
    if plate_elevation is None:
        plate_elevation = elevation
    source = u"Плиты/крыши/фундаменты ({:.2f} м)".format(plate_elevation * FT_TO_M)
    return cells, source, 0, []


def unique_elevations(values, tolerance):
    result = []
    for value in sorted(values):
        if not result or value - result[-1] > tolerance:
            result.append(value)
        else:
            result[-1] = max(result[-1], value)
    return result


def median(values):
    ordered = sorted(values)
    if not ordered:
        return None
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return 0.5 * (ordered[middle - 1] + ordered[middle])


def expand_vertical_edges(edges):
    if len(edges) < 2:
        return edges, 3.0 * M_TO_FT, 0
    gaps = [edges[index + 1] - edges[index] for index in range(len(edges) - 1)]
    normal_gaps = [gap for gap in gaps if 2.20 * M_TO_FT <= gap <= 5.00 * M_TO_FT]
    typical = median(normal_gaps)
    if typical is None:
        typical = 3.0 * M_TO_FT
    expanded = [edges[0]]
    added = 0
    for index in range(len(edges) - 1):
        z0 = edges[index]
        z1 = edges[index + 1]
        gap = z1 - z0
        if gap > max(4.50 * M_TO_FT, typical * 1.60):
            count = max(2, int(math.ceil(gap / typical)))
            step = gap / count
            for part in range(1, count):
                expanded.append(z0 + part * step)
                added += 1
        expanded.append(z1)
    return expanded, typical, added


def level_name_at(elevation, levels):
    tolerance = 0.05 * M_TO_FT
    closest = None
    closest_distance = tolerance
    for level in levels:
        distance = abs(safe_level_elevation(level) - elevation)
        if distance < closest_distance:
            closest = level
            closest_distance = distance
    try:
        return to_unicode(closest.Name) if closest is not None else None
    except Exception:
        return None


def interpolate_triangle_z(x, y, triangle):
    a, b, c = triangle
    denominator = (b.Y - c.Y) * (a.X - c.X) + (c.X - b.X) * (a.Y - c.Y)
    if abs(denominator) < 1e-12:
        return None
    wa = ((b.Y - c.Y) * (x - c.X) + (c.X - b.X) * (y - c.Y)) / denominator
    wb = ((c.Y - a.Y) * (x - c.X) + (a.X - c.X) * (y - c.Y)) / denominator
    wc = 1.0 - wa - wb
    tolerance = -1e-7
    if wa < tolerance or wb < tolerance or wc < tolerance:
        return None
    return wa * a.Z + wb * b.Z + wc * c.Z


def roof_height_map(triangles, grid, footprint, roof_base):
    result = {}
    if not triangles or not footprint:
        return result
    for triangle in triangles:
        xs = [point.X for point in triangle]
        ys = [point.Y for point in triangle]
        i0, j0, i1, j1 = grid.index_bounds(min(xs), min(ys), max(xs), max(ys))
        for j in range(j0, j1 + 1):
            for i in range(i0, i1 + 1):
                cell = (i, j)
                if cell not in footprint:
                    continue
                x, y = grid.center(i, j)
                elevation = interpolate_triangle_z(x, y, triangle)
                if elevation is not None and elevation >= roof_base - 0.05 * M_TO_FT:
                    result[cell] = max(result.get(cell, elevation), elevation)

    missing = set(footprint) - set(result.keys())
    for _ in range(4):
        if not missing:
            break
        additions = {}
        for i, j in missing:
            nearby = []
            for di in (-1, 0, 1):
                for dj in (-1, 0, 1):
                    value = result.get((i + di, j + dj))
                    if value is not None:
                        nearby.append(value)
            if nearby:
                additions[(i, j)] = sum(nearby) / len(nearby)
        result.update(additions)
        missing -= set(additions.keys())

    if missing:
        roof_values = [point.Z for triangle in triangles for point in triangle]
        fallback = 0.5 * (min(roof_values) + max(roof_values)) if roof_values else roof_base
        for cell in missing:
            result[cell] = max(roof_base, fallback)
    return result


def row_runs(indices):
    values = sorted(indices)
    if not values:
        return []
    runs = []
    start = values[0]
    previous = values[0]
    for value in values[1:]:
        if value == previous + 1:
            previous = value
        else:
            runs.append((start, previous))
            start = value
            previous = value
    runs.append((start, previous))
    return runs


def rectangles_from_cells(cells):
    rows = {}
    for i, j in cells:
        rows.setdefault(j, []).append(i)
    active = {}
    rectangles = []
    previous_row = None
    for j in sorted(rows.keys()):
        runs = set(row_runs(rows[j]))
        if previous_row is None or j != previous_row + 1:
            for run, span in active.items():
                rectangles.append((run[0], run[1], span[0], span[1]))
            active = {}
        for run in list(active.keys()):
            if run not in runs:
                span = active.pop(run)
                rectangles.append((run[0], run[1], span[0], span[1]))
        for run in runs:
            if run in active:
                active[run][1] = j
            else:
                active[run] = [j, j]
        previous_row = j
    for run, span in active.items():
        rectangles.append((run[0], run[1], span[0], span[1]))
    return rectangles


def create_rectangle_solid(grid, rectangle, z0, z1):
    i0, i1, j0, j1 = rectangle
    x0 = grid.xmin + i0 * grid.cell
    x1 = grid.xmin + (i1 + 1) * grid.cell
    y0 = grid.ymin + j0 * grid.cell
    y1 = grid.ymin + (j1 + 1) * grid.cell
    points = [
        DB.XYZ(x0, y0, z0),
        DB.XYZ(x1, y0, z0),
        DB.XYZ(x1, y1, z0),
        DB.XYZ(x0, y1, z0),
    ]
    loop = DB.CurveLoop()
    for index in range(4):
        loop.Append(DB.Line.CreateBound(points[index], points[(index + 1) % 4]))
    loops = List[DB.CurveLoop]()
    loops.Add(loop)
    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(loops, DB.XYZ.BasisZ, z1 - z0)


def merge_shape_bands(bands):
    merged = []
    for band in bands:
        if not band[u"cells"] or band[u"z1"] - band[u"z0"] < 1e-8:
            continue
        if merged and abs(merged[-1][u"z1"] - band[u"z0"]) < 1e-7 and merged[-1][u"cells"] == band[u"cells"]:
            merged[-1][u"z1"] = band[u"z1"]
        else:
            merged.append({u"z0": band[u"z0"], u"z1": band[u"z1"], u"cells": band[u"cells"]})
    return merged


def build_shape_solids(grid, bands):
    solids = []
    failed = 0
    for band in merge_shape_bands(bands):
        for rectangle in rectangles_from_cells(band[u"cells"]):
            try:
                solids.append(create_rectangle_solid(grid, rectangle, band[u"z0"], band[u"z1"]))
            except Exception:
                failed += 1
    return solids, failed


def compute_plate_volume(plates, levels, grid, ground_z, plate_zmin):
    groups = build_plate_cell_groups(plates, grid, strict=True)
    groups.sort(key=lambda item: item[0])
    if not groups:
        return None

    rows = []
    shape_bands = []
    total_ft3 = 0.0
    above_ft3 = 0.0
    below_ft3 = 0.0

    expanded_groups = []
    all_cells = set()
    for elevation, cells in groups:
        expanded_groups.append((elevation, cells, cells))
        all_cells.update(cells)

    interval_cells = {}
    for cell in all_cells:
        elevations = []
        for elevation, original_cells, expanded_cells in expanded_groups:
            if cell in expanded_cells:
                elevations.append(elevation)
        elevations = unique_elevations(elevations, 0.05 * M_TO_FT)
        for index in range(len(elevations) - 1):
            z0 = elevations[index]
            z1 = elevations[index + 1]
            if z1 - z0 > MIN_BAND_M * M_TO_FT:
                interval_cells.setdefault((z0, z1), set()).add(cell)

    first_elevation, first_cells = groups[0]
    if first_elevation > plate_zmin + 1e-8:
        interval_cells.setdefault((plate_zmin, first_elevation), set()).update(first_cells)

    band_data = []
    for interval in sorted(interval_cells.keys()):
        z0, z1 = interval
        cells = interval_cells[interval]
        if cells:
            band_data.append((z0, z1, cells, z0, z1))

    for z0, z1, cells, source_elevation, target_elevation in band_data:
        area_ft2 = len(cells) * grid.cell * grid.cell
        height = z1 - z0
        volume = area_ft2 * height
        total_ft3 += volume
        if ground_z is None:
            above_ft3 += volume
        else:
            below_height = max(0.0, min(z1, ground_z) - z0)
            above_height = max(0.0, z1 - max(z0, ground_z))
            below_ft3 += area_ft2 * below_height
            above_ft3 += area_ft2 * above_height
        name0 = level_name_at(z0, levels)
        name1 = level_name_at(z1, levels)
        label = u"{} → {}".format(name0, name1) if name0 and name1 else u"от {:.2f} до {:.2f} м".format(z0 * FT_TO_M, z1 * FT_TO_M)
        rows.append([
            label,
            u"{:.2f}".format(area_ft2 * FT2_TO_M2),
            u"{:.2f}".format(height * FT_TO_M),
            u"{:.2f}".format(volume * FT3_TO_M3),
            u"Вертикальные колонки плит {:.2f} → {:.2f} м".format(
                source_elevation * FT_TO_M,
                target_elevation * FT_TO_M,
            ),
            u"—",
        ])
        shape_bands.append({u"z0": z0, u"z1": z1, u"cells": cells})

    if not shape_bands:
        return None
    gaps = [groups[index + 1][0] - groups[index][0] for index in range(len(groups) - 1)]
    typical = median([gap for gap in gaps if gap > MIN_BAND_M * M_TO_FT])
    if typical is None:
        typical = 0.0
    warnings = []
    return {
        u"rows": rows,
        u"shape_bands": shape_bands,
        u"warnings": warnings,
        u"total_ft3": total_ft3,
        u"above_ft3": above_ft3,
        u"below_ft3": below_ft3,
        u"roof_ft3": 0.0,
        u"zmin": plate_zmin,
        u"standard_top": groups[-1][0],
        u"roof_base_raw": None,
        u"roof_max": None,
        u"inferred_storey_height": typical,
        u"generated_edges": 0,
        u"level_count": 0,
        u"floor_elevation_count": len(groups),
        u"mode": MODE_PLATES,
        u"top_plate_area_ft2": len(groups[-1][1]) * grid.cell * grid.cell,
    }


def compute(elements, walls, floors, roof_triangles, levels, grid, ground_z, mode):
    bounds = [element_bounds(elem) for elem in elements]
    bounds = [item for item in bounds if item is not None]
    if not bounds:
        return None
    zmin = min(item[4] for item in bounds)
    zmax = max(item[5] for item in bounds)

    if mode == MODE_PLATES and floors:
        plate_bounds = []
        seen_ids = set()
        for record in floors:
            record_id = element_id(record.elem)
            if record_id in seen_ids:
                continue
            seen_ids.add(record_id)
            record_bounds = element_bounds(record.elem)
            if record_bounds is not None:
                plate_bounds.append(record_bounds)
        if plate_bounds:
            zmin = min(item[4] for item in plate_bounds)
            zmax = max(item[5] for item in plate_bounds)
        plate_result = compute_plate_volume(floors, levels, grid, ground_z, zmin)
        if plate_result is not None:
            return plate_result

    wall_top_values = [wall.zmax for wall in walls]
    floor_top_values = [record.elevation for record in floors]
    structure_top_candidates = []
    if wall_top_values and mode != MODE_PLATES:
        structure_top_candidates.append(max(wall_top_values))
    if floor_top_values:
        structure_top_candidates.append(max(floor_top_values))
    structure_top = max(structure_top_candidates) if structure_top_candidates else zmax

    if roof_triangles and mode != MODE_PLATES:
        roof_vertices = [point.Z for triangle in roof_triangles for point in triangle]
        roof_base = min(roof_vertices)
        roof_max = max(roof_vertices)
        standard_top = max(zmin, structure_top)
    else:
        roof_base = None
        roof_max = None
        standard_top = max(zmin, structure_top)

    edge_values = [zmin, standard_top]
    if mode != MODE_PLATES:
        for level in levels:
            elevation = safe_level_elevation(level)
            if zmin < elevation < standard_top:
                edge_values.append(elevation)
    for floor in floors:
        if zmin < floor.elevation < standard_top:
            edge_values.append(floor.elevation)
    edges = unique_elevations(edge_values, EDGE_MERGE_M * M_TO_FT)
    if edges[0] > zmin + 1e-8:
        edges.insert(0, zmin)
    if edges[-1] < standard_top - 1e-8:
        edges.append(standard_top)
    edges, inferred_storey_height, generated_edges = expand_vertical_edges(edges)

    rows = []
    shape_bands = []
    warnings = []
    warning_counts = {}
    total_ft3 = 0.0
    above_ft3 = 0.0
    below_ft3 = 0.0
    last_cells = set()

    band_pairs = []
    for index in range(len(edges) - 1):
        z0, z1 = edges[index], edges[index + 1]
        if z1 - z0 >= MIN_BAND_M * M_TO_FT:
            band_pairs.append((z0, z1))

    for index, pair in enumerate(band_pairs):
        z0, z1 = pair
        try:
            output.update_progress(index + 1, max(1, len(band_pairs)))
        except Exception:
            pass
        if mode == MODE_PLATES:
            cells, source, active_walls, band_warnings = choose_plate_footprint(
                floors, grid, z0 + 0.01 * M_TO_FT
            )
        else:
            cells, source, active_walls, band_warnings = choose_footprint(
                walls, floors, grid, 0.5 * (z0 + z1)
            )
        for message in band_warnings:
            warning_counts[message] = warning_counts.get(message, 0) + 1
        if not cells and last_cells:
            cells = last_cells
            source = u"Предыдущий контур (резерв)"
            warning_counts[u"Пустой диапазон заполнен предыдущим контуром"] = warning_counts.get(u"Пустой диапазон заполнен предыдущим контуром", 0) + 1
        if cells:
            last_cells = cells
        area_ft2 = len(cells) * grid.cell * grid.cell
        height = z1 - z0
        volume = area_ft2 * height
        total_ft3 += volume
        if ground_z is None:
            above_ft3 += volume
        else:
            below_height = max(0.0, min(z1, ground_z) - z0)
            above_height = max(0.0, z1 - max(z0, ground_z))
            below_ft3 += area_ft2 * below_height
            above_ft3 += area_ft2 * above_height
        name0 = level_name_at(z0, levels)
        name1 = level_name_at(z1, levels)
        label = u"{} → {}".format(name0, name1) if name0 and name1 else u"от {:.2f} до {:.2f} м".format(z0 * FT_TO_M, z1 * FT_TO_M)
        rows.append([
            label,
            u"{:.2f}".format(area_ft2 * FT2_TO_M2),
            u"{:.2f}".format(height * FT_TO_M),
            u"{:.2f}".format(volume * FT3_TO_M3),
            source,
            active_walls,
        ])
        shape_bands.append({u"z0": z0, u"z1": z1, u"cells": cells})

    roof_volume = 0.0
    if mode != MODE_PLATES and roof_triangles and last_cells and roof_max > standard_top + 1e-8:
        roof_top = roof_max if roof_max is not None else standard_top
        heights = roof_height_map(roof_triangles, grid, last_cells, standard_top)
        for cell in last_cells:
            top = max(standard_top, heights.get(cell, standard_top))
            roof_volume += grid.cell * grid.cell * (top - standard_top)
        total_ft3 += roof_volume
        if ground_z is None or ground_z <= standard_top:
            above_ft3 += roof_volume
        else:
            for cell in last_cells:
                top = max(standard_top, heights.get(cell, standard_top))
                below_ft3 += grid.cell * grid.cell * max(0.0, min(top, ground_z) - standard_top)
                above_ft3 += grid.cell * grid.cell * max(0.0, top - max(standard_top, ground_z))
        rows.append([
            u"Кровельный объём",
            u"{:.2f}".format(len(last_cells) * grid.cell * grid.cell * FT2_TO_M2),
            u"переменная",
            u"{:.2f}".format(roof_volume * FT3_TO_M3),
            u"Высотная карта кровли",
            u"—",
        ])
        roof_step = max(grid.cell, 0.25 * M_TO_FT)
        z0 = standard_top
        while z0 < roof_top - 1e-8:
            z1 = min(roof_top, z0 + roof_step)
            midpoint = 0.5 * (z0 + z1)
            cells = {cell for cell in last_cells if heights.get(cell, standard_top) >= midpoint}
            if cells:
                shape_bands.append({u"z0": z0, u"z1": z1, u"cells": cells})
            z0 = z1

    for message, count in sorted(warning_counts.items()):
        warnings.append([message, count])

    return {
        u"rows": rows,
        u"shape_bands": shape_bands,
        u"warnings": warnings,
        u"total_ft3": total_ft3,
        u"above_ft3": above_ft3,
        u"below_ft3": below_ft3,
        u"roof_ft3": roof_volume,
        u"zmin": zmin,
        u"standard_top": standard_top,
        u"roof_base_raw": roof_base,
        u"roof_max": roof_max,
        u"inferred_storey_height": inferred_storey_height,
        u"generated_edges": generated_edges,
        u"level_count": len([level for level in levels if zmin < safe_level_elevation(level) < standard_top]),
        u"floor_elevation_count": len(unique_elevations(floor_top_values, 0.05 * M_TO_FT)),
        u"mode": mode,
    }


def find_old_mass_ids():
    old_ids = List[DB.ElementId]()
    try:
        direct_shapes = DB.FilteredElementCollector(doc).OfClass(DB.DirectShape).WhereElementIsNotElementType()
        for item in direct_shapes:
            try:
                is_calculated = to_unicode(item.Name).startswith(DS_NAME_PREFIX)
                try:
                    is_calculated = is_calculated or to_unicode(item.ApplicationId) == DS_NAME_PREFIX
                except Exception:
                    pass
                if is_calculated:
                    old_ids.Add(item.Id)
            except Exception:
                pass
    except Exception:
        pass
    return old_ids


def delete_old_masses():
    old_ids = find_old_mass_ids()
    if old_ids.Count == 0:
        return 0
    try:
        deleted_ids = doc.Delete(old_ids)
        return deleted_ids.Count
    except Exception:
        pass
    deleted = 0
    for old_id in old_ids:
        try:
            doc.Delete(old_id)
            deleted += 1
        except Exception:
            pass
    return deleted


def create_direct_shape(solids):
    category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Mass)
    if category is None:
        raise Exception(u"Категория Mass недоступна")
    direct_shape = DB.DirectShape.CreateElement(doc, category.Id)
    try:
        direct_shape.ApplicationId = DS_NAME_PREFIX
        direct_shape.ApplicationDataId = datetime.now().strftime("%Y%m%d_%H%M%S")
    except Exception:
        pass
    shapes = List[DB.GeometryObject]()
    for solid in solids:
        try:
            if direct_shape.IsValidShape(solid):
                shapes.Add(solid)
        except Exception:
            shapes.Add(solid)
    if shapes.Count == 0:
        raise Exception(u"DirectShape не принял ни одного расчётного солида")
    direct_shape.SetShape(shapes)
    try:
        direct_shape.Name = u"{}_{}".format(DS_NAME_PREFIX, datetime.now().strftime("%Y%m%d_%H%M%S"))
    except Exception:
        pass
    return direct_shape, shapes.Count


def unhide_mass_category():
    try:
        view = uidoc.ActiveView
        category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Mass)
        if category is not None:
            view.SetCategoryHidden(category.Id, False)
    except Exception:
        pass


def show_created_element(elem):
    if elem is None:
        return
    try:
        ids = List[DB.ElementId]()
        ids.Add(elem.Id)
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        pass
    try:
        uidoc.ShowElements(elem)
    except Exception:
        pass


def export_csv(calc):
    path = forms.save_file(file_ext=u"csv", default_file_name=u"obem_zdaniya.csv")
    if not path:
        return
    lines = [u"Диапазон;Площадь контура, м²;Высота, м;Объём, м³;Источник;Активных стен"]
    for row in calc[u"rows"]:
        lines.append(u";".join([to_unicode(value) for value in row]))
    lines.append(u"")
    lines.append(u"Итого;;;{:.2f};;".format(calc[u"total_ft3"] * FT3_TO_M3))
    try:
        with io.open(path, u"w", encoding=u"utf-8-sig") as handle:
            handle.write(u"\n".join(lines))
    except Exception as ex:
        forms.alert(u"Не удалось сохранить CSV: {}".format(to_unicode(ex)), title=u"Объём здания")


def print_report(calc, source, grid, walls, floors, roof_triangles, direct_shape, solid_count, failed_shapes, deleted):
    mode_label = u"Гибридный: стены + плиты/крыши/фундаменты"
    if calc[u"mode"] == MODE_PLATES:
        mode_label = u"Упрощённый: только плиты/крыши/фундаменты"
    output.print_md(u"## Строительный объём здания")
    output.print_md(u"> Общий объём: **{:.2f} м³**".format(calc[u"total_ft3"] * FT3_TO_M3))
    output.print_md(u"> Надземный: **{:.2f} м³** | Подземный: **{:.2f} м³**".format(
        calc[u"above_ft3"] * FT3_TO_M3,
        calc[u"below_ft3"] * FT3_TO_M3,
    ))
    output.print_md(u"> Источник: **{}** | Сетка запрошена: **{:.2f} м** | фактическая: **{:.2f} м** ({} × {} ячеек)".format(
        source, grid.requested_cell * FT_TO_M, grid.cell * FT_TO_M, grid.nx, grid.ny
    ))
    output.print_md(u"> Максимальная расчётная погрешность наружной границы растра: **до {:.2f} м**".format(
        grid.cell * FT_TO_M
    ))
    output.print_md(u"> Режим расчёта: **{}**".format(mode_label))
    roof_top_text = u"нет"
    if calc[u"roof_max"] is not None:
        roof_top_text = u"{:.2f} м".format(calc[u"roof_max"] * FT_TO_M)
    output.print_md(u"> Вертикальный диапазон: низ **{:.2f} м** | верх этажной части **{:.2f} м** | верх кровли **{}**".format(
        calc[u"zmin"] * FT_TO_M,
        calc[u"standard_top"] * FT_TO_M,
        roof_top_text,
    ))
    output.print_md(u"> Уровней внутри диапазона: **{}** | отметок перекрытий: **{}** | добавлено промежуточных срезов: **{}** | типовой шаг: **{:.2f} м**".format(
        calc[u"level_count"],
        calc[u"floor_elevation_count"],
        calc[u"generated_edges"],
        calc[u"inferred_storey_height"] * FT_TO_M,
    ))
    if calc[u"mode"] == MODE_PLATES:
        output.print_md(u"> Максимальный верх формы: **{:.2f} м** | площадь самой верхней группы плит: **{:.2f} м²**".format(
            calc[u"standard_top"] * FT_TO_M,
            calc[u"top_plate_area_ft2"] * FT2_TO_M2,
        ))
    output.print_md(u"> Контурообразующих стен: **{}** | Контуров плит/крыш/фундаментов: **{}** | Треугольников кровли: **{}**".format(
        len(walls), len(floors), len(roof_triangles)
    ))
    if direct_shape is not None:
        output.print_md(u"> Проверочный Mass: {} | ID: **{}** | простых призм: **{}**".format(
            output.linkify(direct_shape.Id), direct_shape.Id.IntegerValue, solid_count
        ))
    if deleted:
        output.print_md(u"> Удалены прежние расчётные Mass: **{}**".format(deleted))
    if failed_shapes:
        output.print_md(u"> Не создано проверочных призм: **{}** (на числовой объём не влияет)".format(failed_shapes))

    output.print_md(u"### Расчёт по высотным диапазонам")
    output.print_table(
        table_data=calc[u"rows"],
        columns=[u"Диапазон", u"Площадь, м²", u"Высота, м", u"Объём, м³", u"Источник", u"Стен"],
        formats=[u"{}", u"{}", u"{}", u"{}", u"{}", u"{}"],
    )
    if calc[u"warnings"]:
        output.print_md(u"### Коррекции контура")
        output.print_table(
            table_data=calc[u"warnings"],
            columns=[u"Причина", u"Диапазонов"],
            formats=[u"{}", u"{}"],
        )
    output.print_md(u"---")
    if calc[u"mode"] == MODE_PLATES:
        output.print_md(u"Число рассчитано по локальным вертикальным колонкам сетки. Для каждой XY-ячейки "
                        u"находятся все перекрытия/крыши/фундаментные плиты над ней, и пространство заполняется "
                        u"между каждой последовательной парой. Пропущенная промежуточная плита не создаёт разрыв, "
                        u"а часть здания без следующей плиты сверху останавливается на собственной кровле. "
                        u"Наружный контур не расширяется морфологически и не получает допуск соседней ячейки. "
                        u"Проверочный Mass имеет ступенчатый контур с шагом указанной сетки.")
    else:
        output.print_md(u"Число рассчитано по заполненным ячейкам внутри наружного контура стен, "
                        u"дополненного перекрытиями и крышами. Окна, двери, помещения, внутренние стены "
                        u"и шахты не вычитаются.")


def main():
    elements, source = get_relevant_elements()
    if not elements:
        forms.alert(u"Не найдены стены, перекрытия, кровли или фундаменты для расчёта.", title=u"Объём здания")
        return

    mode_choice = ask_calculation_mode()
    if mode_choice is None:
        return

    levels = get_levels()
    ground_level, accepted = ask_ground_level(levels)
    if not accepted:
        return
    cell_m = ask_cell_size()
    if cell_m is None:
        return

    options = geometry_options()
    walls, exterior_function_used = build_wall_records(elements)
    floors = build_floor_records(elements, options)
    roof_triangles = []
    if mode_choice.key != MODE_PLATES:
        roof_triangles = build_roof_triangles(elements, options)
    if not walls and not floors:
        forms.alert(u"Не удалось получить оси стен или наружные контуры перекрытий.", title=u"Объём здания")
        return
    if mode_choice.key == MODE_PLATES and not floors:
        forms.alert(u"Для упрощённого режима не найдены контуры перекрытий, крыш или фундаментных плит.", title=u"Объём здания")
        return

    bounds = [element_bounds(elem) for elem in elements]
    bounds = [item for item in bounds if item is not None]
    if not bounds:
        forms.alert(u"Не удалось определить габариты выбранных элементов.", title=u"Объём здания")
        return
    margin = GRID_MARGIN_M * M_TO_FT
    xmin = min(item[0] for item in bounds) - margin
    ymin = min(item[1] for item in bounds) - margin
    xmax = max(item[2] for item in bounds) + margin
    ymax = max(item[3] for item in bounds) + margin
    grid = Grid(xmin, ymin, xmax, ymax, cell_m * M_TO_FT)

    ground_z = safe_level_elevation(ground_level) if ground_level is not None else None
    calc = compute(elements, walls, floors, roof_triangles, levels, grid, ground_z, mode_choice.key)
    if calc is None or calc[u"total_ft3"] <= 1e-8:
        forms.alert(u"Расчётный внешний объём не получен. Проверьте замкнутость наружных стен и наличие перекрытий.", title=u"Объём здания")
        return

    solids, failed_shapes = build_shape_solids(grid, calc[u"shape_bands"])
    if not solids:
        forms.alert(u"Объём рассчитан, но проверочную форму создать не удалось.", title=u"Объём здания")
        return

    direct_shape = None
    solid_count = 0
    deleted = 0

    cleanup = DB.Transaction(doc, u"Объём здания: удалить прежние Mass")
    cleanup.Start()
    try:
        deleted = delete_old_masses()
        cleanup.Commit()
    except Exception:
        try:
            if cleanup.GetStatus() == DB.TransactionStatus.Started:
                cleanup.RollBack()
        except Exception:
            pass
        output.print_md(u"> Не удалось удалить прежние расчётные Mass (некритично)")
        output.print_md(u"```\n{}\n```".format(to_unicode(traceback.format_exc())))

    transaction = DB.Transaction(doc, u"Объём здания: создать проверочный Mass")
    transaction.Start()
    try:
        direct_shape, solid_count = create_direct_shape(solids)
        unhide_mass_category()
        transaction.Commit()
    except Exception as ex:
        try:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
        except Exception:
            pass
        output.print_md(u"### Ошибка создания проверочного Mass")
        output.print_md(u"```\n{}\n```".format(to_unicode(traceback.format_exc())))
        return

    show_created_element(direct_shape)
    visibility_transaction = DB.Transaction(doc, u"Объём здания: показать Mass")
    visibility_transaction.Start()
    try:
        unhide_mass_category()
        visibility_transaction.Commit()
    except Exception:
        try:
            if visibility_transaction.GetStatus() == DB.TransactionStatus.Started:
                visibility_transaction.RollBack()
        except Exception:
            pass
    print_report(calc, source, grid, walls, floors, roof_triangles, direct_shape, solid_count, failed_shapes, deleted)

    try:
        if forms.alert(u"Сохранить отчёт в CSV?", title=u"Объём здания", ok=False, yes=True, no=True):
            export_csv(calc)
    except Exception:
        pass

    if mode_choice.key == MODE_PLATES:
        contour_mode = u"плиты/крыши/фундаменты"
    else:
        contour_mode = u"наружные стены" if exterior_function_used else u"все выбранные стены"
    forms.alert(
        u"Общий объём: {:.2f} м³\nMass ID: {}\nКонтур: {}\nЭлемент выделен в модели.".format(
            calc[u"total_ft3"] * FT3_TO_M3,
            direct_shape.Id.IntegerValue,
            contour_mode,
        ),
        title=u"Объём здания"
    )


if __name__ == u"__main__":
    main()
