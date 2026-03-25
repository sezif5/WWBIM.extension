# Task 4 QA Checklist: Verify fill_dimensions.py fills ADSK_Размер_Объём

## Prerequisites

1. **Ensure Revit uses the worktree copy** of `fill_dimensions.py`:
   - Worktree location: `D:\GitHub\WWBIM.extension\.worktrees\fill-dimensions-revit2022-volume\WWBIM.extension\lib\Batch Operations\fill_dimensions.py`
   - **Option A**: If Revit loads pyRevit from the worktree folder directly, you're set.
   - **Option B**: Temporarily copy `fill_dimensions.py` from the worktree to the deployed extension location that Revit uses (find via `pyrevit paths` command or Revit journal files).
   - **Do NOT** run with the old repo copy - it will crash on import.

2. **Test model requirements**:
   - At least 1 Structural Column (`OST_StructuralColumns`)
   - At least 1 Structural Framing element (`OST_StructuralFraming`)
   - Elements should have non-zero volume (avoid tiny/nominal elements)

---

## Execution Option 1: Batch Runner UI (Preferred)

### Step 1: Run the script via batch runner

1. Open Revit 2022 with a test model meeting prerequisites.
2. Click `BIM` tab > `Пакетные операции` button.
3. Select "Выполнение python скриптов из библиотеки".
4. Choose `fill_dimensions.py` from the list.
5. Run the operation on the current model.

### Step 2: Verify no import-time error

**Expected result**: Script loads and executes without `AttributeError: type object 'BuiltInParameter' has no attribute '...'`

**Capture evidence**:
```batch
dir /b .sisyphus\evidence\
type nul > .sisyphus\evidence\task-4-run-ok.txt
```

Then paste the batch runner console output into `.sisyphus\evidence\task-4-run-ok.txt`.
If successful, you should see execution completion with counters like `updated_elements: N`.

---

## Execution Option 2: RevitPythonShell/pyRevit Console (Alternative)

### Step 1: Open RevitPythonShell console

Press `F8` in Revit 2022 to open the console.

### Step 2: Run the script and capture two sample elements

Paste this snippet into the console:

```python
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
import os

# Create evidence directory
evidence_dir = r"D:\GitHub\WWBIM.extension\.sisyphus\evidence"
os.makedirs(evidence_dir, exist_ok=True)

# Import and execute fill_dimensions
import sys
script_path = r"D:\GitHub\WWBIM.extension\.worktrees\fill-dimensions-revit2022-volume\WWBIM.extension\lib\Batch Operations\fill_dimensions.py"
module_name = "fill_dimensions_qa"

# Clear cache to ensure fresh import
sys.modules.pop(module_name, None)

import imp
module = imp.load_source(module_name, script_path)

# Execute on active document
result = module.Execute(__revit__.ActiveUIDocument.Document)

# Capture execution result to file
run_ok_path = os.path.join(evidence_dir, "task-4-run-ok.txt")
with open(run_ok_path, "w") as f:
    f.write("Execute result: {}\n".format(result))

# Get one element from each structural category
doc = __revit__.ActiveUIDocument.Document
structural_columns = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToElements())
structural_framing = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements())

sample_col = structural_columns[0] if structural_columns else None
sample_frame = structural_framing[0] if structural_framing else None

# Read ADSK_Размер_Объём for both
volume_filled_path = os.path.join(evidence_dir, "task-4-volume-filled.txt")
with open(volume_filled_path, "w") as f:
    if sample_col:
        param = sample_col.LookupParameter("ADSK_Размер_Объём")
        val = param.AsString() if param else "None"
        cat = sample_col.Category.Name if sample_col.Category else "Unknown"
        f.write("Category: {}, ElementId: {}, ADSK_Размер_Объём: {}\n".format(cat, sample_col.Id.IntegerValue, val))
    else:
        f.write("No Structural Columns found in model.\n")

    if sample_frame:
        param = sample_frame.LookupParameter("ADSK_Размер_Объём")
        val = param.AsString() if param else "None"
        cat = sample_frame.Category.Name if sample_frame.Category else "Unknown"
        f.write("Category: {}, ElementId: {}, ADSK_Размер_Объём: {}\n".format(cat, sample_frame.Id.IntegerValue, val))
    else:
        f.write("No Structural Framing found in model.\n")

print("QA evidence saved to: {}".format(evidence_dir))
```

### Step 3: Check console output

**Expected result**:
- No `AttributeError` during import or execution
- Console prints `QA evidence saved to: ...`
- Check `.sisyphus\evidence\task-4-run-ok.txt` shows `Execute result: {'success': True, ...}`

---

## Step 4: Verify ADSK_Размер_Объём values

Open `.sisyphus\evidence\task-4-volume-filled.txt` and check:

1. **Both elements found**: One Structural Column and one Structural Framing element are listed.
2. **Values non-empty**: `ADSK_Размер_Объём` is not `None` or empty string for both.
3. **Values parseable**: For known-nonzero elements, the value should parse as a float > 0.

**Example of successful output**:
```
Category: Structural Columns, ElementId: 12345, ADSK_Размер_Объём: 0.85
Category: Structural Framing, ElementId: 67890, ADSK_Размер_Объём: 0.42
```

**Manual validation** (optional but recommended):
1. Select the two elements in Revit UI.
2. Open Properties dialog.
3. Locate `ADSK_Размер_Объём` parameter.
4. Confirm the value matches what was logged.

---

## Success Criteria

- [ ] No import-time `AttributeError` in Revit 2022
- [ ] Script executes and returns `success: True`
- [ ] `.sisyphus\evidence\task-4-run-ok.txt` exists with execution output
- [ ] `.sisyphus\evidence\task-4-volume-filled.txt` exists with two elements
- [ ] Both elements have non-empty `ADSK_Размер_Объём` values
- [ ] Values are numeric and > 0 for known-nonzero elements

---

## Troubleshooting

**Issue**: Revit loads old repo copy instead of worktree
- Run `pyrevit paths` in RevitPythonShell console to see which folder pyRevit loads from.
- Either configure pyRevit to use the worktree folder, or copy `fill_dimensions.py` to the deployed location.

**Issue**: `AttributeError: type object 'BuiltInParameter' has no attribute '...'`
- You're still using the old repo version. Verify Revit is loading from the worktree.

**Issue**: Elements have zero volume
- Ensure test elements have actual geometry (not tiny/nominal family instances).
- Check that `HOST_VOLUME_COMPUTED` parameter exists and has values in Revit UI for those elements.

**Issue**: `ADSK_Размер_Объём` is None/empty
- Check the parameter is bound to the model. If missing, `fill_dimensions.py` will add it automatically.
- Verify the element category is in the DIMENSION_PARAMETERS sources list (lines 46-51 of fill_dimensions.py).
