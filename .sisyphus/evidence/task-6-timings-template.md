# Task 6: Performance Timing Evidence Template

## Purpose
Record performance timings for batch runner execution to detect regressions.

---

## Timing Procedure

### Step 1: Run baseline
Execute `fill_dimensions.py` on a known-good model and record:
- Total elements count
- Total time (- Per-parameter stats

### Step 2: Run on test model
1. Open Revit 2022 with test model.
2. Run `BIM` > `Пакетные операции`.
3. Select `fill_dimensions.py`.
4. Record start time.
5. Execute.

### Step 2: Record timings
After completion, add timing data to to log:

```
=== FILL Dimensions Performance ===
Date: 
Model: 
Start Time: 
End Time: 
Duration: 

### Step 1: Total elements processed
- Structural Columns: 
- Structural Framing: 
- Other categories: 

### Step 2: Per-parameter stats
| Parameter | Updated | Skipped | No Value |
|---|---|---|
|---|---|---|
|---|---|---|
|---|---|---|
|---|           |---|---|---|---|---|
| ADSK_Размер_Длина | | | | | | | | | | |
| ADSK_Размер_Ширина | | | | | | | | | | |
| ADSK_Размер_Высота | | | | | | | | | | |
| ADSK_Размер_Толщина | | | | | | | | | | |
| ADSK_Площадь | | | | | | | | | | |
| ADSK_Размер_Диаметр | | | | | | | | | | |

---

## Evidence File Location
Save this file as: `.sisyphus/evidence/task-6-timings-<timestamp>.md`

---

## Success Criteria
- Timings recorded for each parameter
- No significant slowdown (> +10% threshold)
- Script completes without exception

---

*Generated: 2026-03-23*
*Plan reference: .sisyphus/plans/fill-dimensions-revit2022-volume-attrerror.md*
