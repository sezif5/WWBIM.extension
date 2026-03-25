# Learnings

- Batch runner loads scripts via `imp.load_source(module_name, script_path)` and then calls `module.Execute(doc)`; see `WWBIM.extension/WW.BIM.tab/BIM.panel/Пакетные операции.pushbutton/script.py:954-970`.
- It clears module cache first: `sys.modules.pop(module_name, None)` at `WWBIM.extension/WW.BIM.tab/BIM.panel/Пакетные операции.pushbutton/script.py:964`.
- `fill_dimensions.py` evaluates `DIMENSION_PARAMETERS` at import time; direct enum access there can crash import before `Execute`.
- Current crash candidate: `BuiltInParameter.STRUCTURAL_VOLUME` in `WWBIM.extension/lib/Batch Operations/fill_dimensions.py:46`.
- Secondary crash risk: `BuiltInParameter.STRUCTURAL_AREA` in `WWBIM.extension/lib/Batch Operations/fill_dimensions.py:287`.
- Repo already has a safe runtime enum resolution pattern using `Enum.IsDefined` + `Enum.Parse`; reference: `WWBIM.extension/lib/Batch Operations/auto_navis_export_script.py:_resolve_bip/_resolve_bic`.
## BuiltInParameter Investigation - Revit 2022 Compatibility

### Files with BuiltInParameter Usage

#### Primary File: fill_dimensions.py
Location: D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_dimensions.py

### Suspicious BuiltInParameter Members (Volume/Area Related)

| Line | Expression | Category | Assessment |
|------|------------|----------|------------|
| 36 | `BuiltInParameter.HOST_VOLUME_COMPUTED` | Host elements (Walls, Floors, Roofs, Ceilings) | SAFE - standard host parameter |
| 46 | `BuiltInParameter.STRUCTURAL_VOLUME` | Structural Columns, Framing | **HIGH RISK** - likely missing in Revit 2022 |
| 53 | `BuiltInParameter.RBS_PIPE_VOLUME_PARAM` | Pipe curves (MEP) | SAFE - standard MEP parameter |
| 60 | `BuiltInParameter.RBS_DUCT_VOLUME_PARAM` | Duct curves (MEP) | SAFE - standard MEP parameter |
| 278 | `BuiltInParameter.HOST_AREA_COMPUTED` | Host elements (Walls, Floors, Roofs, Ceilings) | SAFE - standard host parameter |
| 287 | `BuiltInParameter.STRUCTURAL_AREA` | Structural Columns, Framing | **HIGH RISK** - likely missing in Revit 2022 |

### Other BuiltInParameter Members in fill_dimensions.py (Reference)

| Line | Expression | Category | Assessment |
|------|------------|----------|------------|
| 73 | `CURVE_ELEM_LENGTH` | Length | SAFE |
| 82 | `STRUCTURAL_FRAME_CUT_LENGTH` | Structural framing length | SAFE (likely exists) |
| 88 | `INSTANCE_LENGTH_PARAM` | Length | SAFE |
| 95 | `RBS_PIPE_LENGTH_PARAM` | Pipes (MEP) | SAFE |
| 103 | `RBS_DUCT_LENGTH_PARAM` | Ducts (MEP) | SAFE |
| 111 | `RBS_CABLETRAY_LENGTH_PARAM` | Cable trays (MEP) | SAFE |
| 117 | `RBS_CONDUIT_LENGTH_PARAM` | Conduits (MEP) | SAFE |
| 129 | `WALL_USER_WIDTH_PARAM` | Walls width | SAFE |
| 135 | `FLOOR_ATTR_THICKNESS_PARAM` | Floors thickness | SAFE |
| 141 | `CEILING_THICKNESS` | Ceilings thickness | SAFE |
| 147 | `STRUCTURAL_FOUNDATION_THICKNESS` | Foundations thickness | SAFE (likely exists) |
| 153 | `FAMILY_WIDTH_PARAM` | Families width | SAFE |
| 173 | `RBS_DUCT_WIDTH_PARAM` | Ducts width (MEP) | SAFE |
| 182 | `RBS_CABLETRAY_WIDTH_PARAM` | Cable trays width (MEP) | SAFE |
| 194 | `FAMILY_HEIGHT_PARAM` | Families height | SAFE |
| 215 | `RBS_DUCT_HEIGHT_PARAM` | Ducts height (MEP) | SAFE |
| 224 | `RBS_CABLETRAY_HEIGHT_PARAM` | Cable trays height (MEP) | SAFE |
| 260 | `FAMILY_DEPTH_PARAM` | Families depth | SAFE |
| 300 | `RBS_PIPE_DIAMETER_PARAM` | Pipes diameter (MEP) | SAFE |
| 309 | `RBS_PIPE_OUTER_DIAMETER_PARAM` | Pipes outer diameter (MEP) | SAFE |
| 318 | `RBS_DUCT_DIAMETER_PARAM` | Ducts diameter (MEP) | SAFE |

### Other Files with BuiltInParameter Usage

#### assign_links_to_worksets_script.py
| Line | Expression | Assessment |
|------|------------|------------|
| 164 | `BuiltInParameter.ELEM_PARTITION_PARAM` | SAFE - standard partition parameter |

### Root Cause Analysis

The crash in `fill_dimensions.py` on startup in Revit 2022 is most likely caused by:

1. **Line 46**: `BuiltInParameter.STRUCTURAL_VOLUME` - This enum member appears to be missing or renamed in Revit 2022 API
2. **Line 287**: `BuiltInParameter.STRUCTURAL_AREA` - This enum member might also be missing in Revit 2022 API

### Recommendations

1. **Immediate Fix**: Add try-catch handling around the DIMENSION_PARAMETERS list definition to gracefully handle missing BuiltInParameter members
2. **Version Detection**: Check Revit version at runtime and use alternative parameters for Revit 2022+
3. **Fallback Strategy**: For structural elements, consider:
   - Calculating volume from geometry (Element.Geometry property)
   - Using alternative built-in parameters like `VOLUME` if available
   - Using family instance parameters

### Gotchas

- The DIMENSION_PARAMETERS list is defined at module load time (line 30-328), so the error occurs during import, not during function execution
- This means the module-level reference to `BuiltInParameter.STRUCTURAL_VOLUME` triggers the AttributeError before any try-catch can catch it
- Solution: Move parameter definitions inside a function or use lazy initialization

### Potential Alternative Parameters to Investigate

For Revit 2022+, the following alternatives might exist:
- `ELEMENT_VOLUME` (if available)
- `GEOMETRY_VOLUME` (if available)
- Calculation from `Element.get_Geometry()` and `Solid.Volume`


## Task 2 Findings: Safe Revit Enum Resolution Patterns

### Key Reference: auto_navis_export_script.py

**File**: `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\auto_navis_export_script.py`

#### Pattern 1: Safe Enum Resolution (PRIMARY RECOMMENDATION)

**Location**: Lines 532-552

```python
from System import Enum

def _resolve_bic(name):
    """Вернуть BuiltInCategory по строке или None, если такого имени нет в текущей версии Revit."""
    if not name:
        return None
    try:
        if Enum.IsDefined(BuiltInCategory, name):
            return Enum.Parse(BuiltInCategory, name)
    except Exception:
        pass
    return None


def _resolve_bip(name):
    if not name:
        return None
    try:
        if Enum.IsDefined(BuiltInParameter, name):
            return Enum.Parse(BuiltInParameter, name)
    except Exception:
        pass
    return None
```

**Why this is safe for Revit 2022**:
- Enum access happens at **runtime**, not import time
- `Enum.IsDefined` checks if the enum member exists before parsing
- Returns `None` gracefully if member doesn't exist in the Revit version
- Wrapped in try/except for additional safety
- No import-time AttributeError possible

#### Pattern 2: String-based Configuration + Runtime Resolution

**Location**: Lines 637-658

```python
names = [
    "OST_RvtLinks",
    "OST_LinkInstances",
    "OST_ExportLayer",
    "OST_ImportInstance",
    # ... more string names
]

def _hide_categories_by_names(doc, view, names):
    ids = List[ElementId]()
    for nm in names or []:
        bic = _resolve_bic(nm)  # Resolve string to enum at runtime
        eid = _cat_id(doc, bic)
        if eid:
            ids.Add(eid)
```

**Why this is safe**:
- Enum names stored as **strings** in module-level configuration
- No direct `BuiltInCategory.SOMETHING` access at import time
- Resolution happens during function execution (runtime)
- Missing enum members simply get skipped (`if eid:` check)

#### Pattern 3: Safe BIP Access with Graceful Degradation

**Location**: Lines 555-566

```python
def _try_set_bip_int(element, bip_name, value):
    bip = _resolve_bip(bip_name)
    if bip is None:  # Safe: skip if BIP doesn't exist in this Revit version
        return False
    try:
        p = element.get_Parameter(bip)
        if p and (not p.IsReadOnly):
            p.Set(int(value))
            return True
    except Exception:
        pass
    return False
```

**Usage**: Lines 665-674
```python
_try_set_bip_int(view, "VIEW_SHOW_IMPORT_CATEGORIES", 0)
_try_set_bip_int(view, "VIEW_SHOW_IMPORT_CATEGORIES_IN_VIEW", 0)
_try_set_bip_int(view, "VIEW_SHOW_ANNOTATION_CATEGORIES", 0)
```

**Why this is safe**:
- Uses `_resolve_bip` to safely convert string name to enum
- Returns `False` if BIP doesn't exist (not an error)
- Caller can handle the failure gracefully
- No crash on import time or runtime

### Why fill_dimensions.py Current Pattern is Unsafe

**Current code** (Lines 36-46):
```python
"SOURCES": [
    {
        "BIP": BuiltInParameter.HOST_VOLUME_COMPUTED,  # Safe in 2022
        "CATEGORIES": [...],
    },
    {
        "BIP": BuiltInParameter.STRUCTURAL_VOLUME,  # CRASHES in Revit 2022!
        "CATEGORIES": [
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming,
        ],
    },
]
```

**Problem**: Direct enum access at module level happens at **import time**, causing AttributeError if the enum member doesn't exist in the running Revit version.

### Recommended Pattern for fill_dimensions.py

**Step 1**: Store BIP names as strings in DIMENSION_PARAMETERS:
```python
DIMENSION_PARAMETERS = [
    {
        "NAME": "ADSK_Размер_Объём",
        "UNIT_TYPE": "volume",
        "SOURCES": [
            {
                "BIP_NAME": "HOST_VOLUME_COMPUTED",  # String, not enum
                "CATEGORIES": [
                    "OST_Walls",  # Also use strings for categories
                    "OST_Floors",
                    # ...
                ],
            },
            {
                "BIP_NAME": "STRUCTURAL_VOLUME",  # Will resolve to None in Revit 2022
                "CATEGORIES": [
                    "OST_StructuralColumns",
                    "OST_StructuralFraming",
                ],
            },
        ],
    },
]
```

**Step 2**: Add helper functions (from auto_navis_export_script.py):
```python
from System import Enum

def _resolve_bip(name):
    if not name:
        return None
    try:
        if Enum.IsDefined(BuiltInParameter, name):
            return Enum.Parse(BuiltInParameter, name)
    except Exception:
        pass
    return None

def _resolve_bic(name):
    if not name:
        return None
    try:
        if Enum.IsDefined(BuiltInCategory, name):
            return Enum.Parse(BuiltInCategory, name)
    except Exception:
        pass
    return None
```

**Step 3**: Update `_BuildSourceMapForParam` to resolve strings at runtime:
```python
def _BuildSourceMapForParam(param_config):
    param_name = param_config.get("NAME")
    if not param_name:
        return {}
    if param_name in _SOURCE_CACHE:
        return _SOURCE_CACHE[param_name]

    source_map = {}
    for source in param_config.get("SOURCES", []):
        bip_name = source.get("BIP_NAME")  # String name
        if not bip_name:
            continue
        
        # Resolve at runtime - safe for Revit 2022
        bip = _resolve_bip(bip_name)
        if bip is None:  # Skip if BIP doesn't exist in this version
            continue
        
        for cat_name in source.get("CATEGORIES", []):
            # Also resolve category at runtime
            bic = _resolve_bic(cat_name)
            if bic is None:
                continue
            
            if bic not in source_map:
                source_map[bic] = []
            source_map[bic].append(bip)

    _SOURCE_CACHE[param_name] = source_map
    return source_map
```

### Additional Safe Patterns to Consider

The repo also uses `getattr` for optional attributes (line 632):
```python
vtid = getattr(view, "ViewTemplateId", None)
```

However, this doesn't help with enum members because they are accessed as `Enum.MEMBER`, not `obj.attribute`.

### Summary

**Best Pattern to Copy**: `_resolve_bip` and `_resolve_bic` from `auto_navis_export_script.py` (lines 532-552)

**Key Principles**:
1. Store enum names as **strings** in configuration
2. Resolve strings to enums at **runtime** using `Enum.IsDefined` + `Enum.Parse`
3. Return `None` if enum doesn't exist in current Revit version
4. Skip `None` BIPs/categories in processing (graceful degradation)
5. Import `from System import Enum` at module level (safe, it's a .NET class)

**Why this avoids Revit 2022 AttributeError**:
- No direct `BuiltInParameter.STRUCTURAL_VOLUME` access at import time
- If enum member doesn't exist, `_resolve_bip` returns `None`
- Processing code skips `None` BIPs instead of crashing
- Module imports successfully; execution handles missing BIPs gracefully

## Task 1 Findings: Script Path and Failing Attribute

### Batch Runner Control Flow

**Entry Point:**
- File: `D:\GitHub\WWBIM.extension\WWBIM.extension\WW.BIM.tab\BIM.panel\Пакетные операции.pushbutton\script.py`
- Line 1392-1393: `if __name__ == "__main__": main()`

**Path Construction for Scripts:**
- Line 31: `PYTHON_SCRIPTS_DIR = os.path.join(LIB_DIR, "Batch Operations")`
  - Where `LIB_DIR = os.path.join(EXTENSION_ROOT, "lib")` (line 29)
  - Final path: `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\`

**Script Selection Flow:**
1. Line 1281-1307: User selects "Выполнение python скриптов из библиотеки" or "Выполнить python скрипты в открытых документах"
2. Line 1285: `scripts = list_python_scripts()` - scans PYTHON_SCRIPTS_DIR for .py files
3. Line 1294-1301: User selects one or more scripts via UI

**Module Loading (CRITICAL - where error occurs):**
- In `action_run_python_script()` (line 860-1207) or `action_run_python_script_on_opened_docs()` (line 696-857)
- For each script:
  - Line 954/719: `script_path = os.path.join(PYTHON_SCRIPTS_DIR, script_rel_path)`
    - Example: `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_dimensions.py`
  - Line 957/722-726: Module name constructed from relative path (replacing / and \ with _)
  - Line 964/729: `sys.modules.pop(module_name, None)` - clear any previous cached module
  - **Line 966/731: `module = imp.load_source(module_name, script_path)`** ← IMPORT HAPPENS HERE
  - Line 968/733-734: Check for Execute function and call `result = module.Execute(doc)`

### AttributeError Location

**File:** `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_dimensions.py`

**Why it fails at import time:**
- Line 18: `from Autodesk.Revit.DB import BuiltInParameter`
- Line 30-328: `DIMENSION_PARAMETERS` is a module-level constant defined at IMPORT TIME
- Line 46: `"BIP": BuiltInParameter.STRUCTURAL_VOLUME,` ← This line executes when module is imported
- In Revit 2022, `BuiltInParameter.STRUCTURAL_VOLUME` does not exist
- Error: `AttributeError: type object 'BuiltInParameter' has no attribute 'STRUCTURAL_VOLUME'`

**Critical Gotcha:**
The AttributeError occurs at line 731/966 of the batch runner (`imp.load_source()`), NOT at line 734/969 (`module.Execute(doc)`). The script's `Execute()` function is never reached because the import itself fails.

### Control Flow Trace

```
User clicks "Пакетные операции" button
  ↓
script.py main() (line 1259)
  ↓
User selects "Выполнение python скриптов из библиотеки" (line 1281)
  ↓
list_python_scripts() scans PYTHON_SCRIPTS_DIR (line 1285)
  ↓
User selects "fill_dimensions.py" (line 1294)
  ↓
User selects models from txt file (line 1353)
  ↓
action_run_python_script(selected_models, selected_scripts) (line 1385)
  ↓
For each script:
  script_path = os.path.join(PYTHON_SCRIPTS_DIR, "fill_dimensions.py")
    → "D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_dimensions.py"
  ↓
  module = imp.load_source("fill_dimensions.py", script_path)  ← CRASH HERE
    → Python imports the module
    → Line 18: from Autodesk.Revit.DB import BuiltInParameter
    → Line 30: DIMENSION_PARAMETERS = [ ... ]  ← Evaluated at import
    → Line 46: BuiltInParameter.STRUCTURAL_VOLUME  ← AttributeError in Revit 2022
  ↓
  (Never reaches: module.Execute(doc))
```

### Summary

- **Batch runner file:** `D:\GitHub\WWBIM.extension\WWBIM.extension\WW.BIM.tab\BIM.panel\Пакетные операции.pushbutton\script.py`
- **Script loading line:** 731 (action_run_python_script_on_opened_docs) or 966 (action_run_python_script)
- **Execute() call line:** 734 or 969 (never reached)
- **Filling script file:** `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_dimensions.py`
- **Failing attribute:** `BuiltInParameter.STRUCTURAL_VOLUME` (line 46)
- **Error occurs:** At import time (module evaluation), not at Execute() time
- **Root cause:** BuiltInParameter.STRUCTURAL_VOLUME doesn't exist in Revit 2022


---
## Research: Volume Parameter for Structural Elements in Revit 2022

**Date**: 2026-03-15

### Finding: No BuiltInParameter for Volume on OST_StructuralColumns/OST_StructuralFraming

**Conclusion**: `HOST_VOLUME_COMPUTED` does NOT apply to structural columns/framing in Revit 2022.

**Evidence**:
- Source: Autodesk Revit API Developer's Guide - Material Quantities
- URL: https://help.autodesk.com/cloudhelp/2018/ENU/Revit-API/Revit_API_Developers_Guide/Revit_Geometric_Elements/Material/Material_quantities.html

**Key Quote**:
> "The methods apply to categories of elements where Category.HasMaterialQuantities property is true. In practice, this is limited to elements that use compound structure, like walls, roofs, floors, ceilings, a few other basic 3D elements like stairs, plus 3D families where materials can be assigned to geometry of family, like windows, doors, **columns**, MEP equipment and fixtures, and generic model families."

### Correct API Method for Revit 2022

For `OST_StructuralColumns` and `OST_StructuralFraming`, use:

**`Element.GetMaterialVolume(ElementId materialId)`**

- Returns volume of specific material for the element
- Since: Revit 2014
- Works for families where materials can be assigned to geometry

**Source**: https://www.revitapidocs.com/2022/99b50d87-bfa6-ca67-e205-47b22cad6587.htm

### Why STRUCTURAL_VOLUME Doesn't Exist

1. **Not in Revit 2022**: Searched Revitapidocs.com for Revit 2022 - no such BuiltInParameter member
2. **Not in Revit 2024/2026**: Checked latest API docs - no such member exists
3. **Typo found in error**: User reported `Structutral_VOLUME` (with typo) - this doesn't exist in any version
4. **No references in codebase**: GitHub searches show no usage of this BuiltInParameter

### Safe Fallback Chain

```python
# Try GetMaterialVolume first (fastest)
material_ids = element.GetMaterialIds()
if material_ids:
    total_volume = 0.0
    for mat_id in material_ids:
        total_volume += element.GetMaterialVolume(mat_id)
else:
    # Fallback to geometry calculation (slower)
    options = revitapp.Create.NewGeometryOptions()
    geom_elem = element.get_Geometry(options)
    for geom_obj in geom_elem:
        if isinstance(geom_obj, Solid):
            total_volume += geom_obj.Volume
```

**Performance Notes**:
- GetMaterialVolume() uses Revit's pre-computed material quantities (fast)
- Geometry calculation is slower but works when material quantification not available
- Always check if material_ids list is empty before iterating

## Task 5 Findings: Documentation Note about Revit-Version-Safe BIPs

### Location
- File: `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\README.md`
- Section: "Особенности реализации" (Implementation Features)
- Position: Added as new subsection #5 after existing 4 points

### Content Added
Four-point subsection about BuiltInParameter safety for Revit version compatibility:

1. **Batch runner mechanism**: Uses `imp.load_source()` - import-time exceptions abort execution
2. **Avoidance warning**: Don't use `BuiltInParameter.X` in module-level constants
3. **Recommended pattern**: Store enum names as strings, resolve at runtime via `Enum.IsDefined` + `Enum.Parse`
4. **Reference example**: Safe pattern implemented in `fill_dimensions.py`

### Documentation Approach
- Minimal scope (4 bullets, within 3-6 limit)
- Placed near `fill_dimensions.py` section for context
- Matches existing README style (Russian language, consistent formatting)
- No code changes, just documentation addition
