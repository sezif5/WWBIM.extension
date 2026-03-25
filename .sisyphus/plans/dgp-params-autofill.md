# DGP Parameters Autofill (Category + Group Model)

## TL;DR
Implement `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/DGP_script.py` to batch-fill instance *Text* shared parameters across the active document using deterministic rules keyed by (Category + optional `Группа модели`). Source parameters can be numeric or text; values are written as project display strings without unit suffixes; write only when changed; missing targets are reported and skipped.

**Deliverables**
- Working pyRevit script: `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/DGP_script.py`
- Updated button metadata to match behavior: `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/bundle.yaml`
- No automated tests (per request); verification is via deterministic report output + Revit run.

**Estimated Effort**: Medium
**Parallel Execution**: YES (3 waves)
**Critical Path**: Helpers -> Core script orchestration -> Category/group rules -> bundle.yaml update

---

## Context

### Original Request (RU)
- "Нужно написать скрипт который в зависимости от Категории и содержания параметра типа Группа модели будет заполнять общие параметры экземпляра. Параметр который заполняем всегда Текст, но параметр исходник может быть числовой."

### Repo Reality
- Target script is currently empty: `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/DGP_script.py`
- Button tooltip currently describes a different feature (kubik sync) and must be updated:
  - `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/bundle.yaml`

### Existing Patterns to Reuse
- Param storage-type helpers: `WWBIM.extension/WW.BIM.tab/Общее.panel/Кубики Синхронизация.pushbutton/Sync_script.py`
- Display formatting via `AsValueString()`: `WWBIM.extension/WW.BIM.tab/BIM.panel/Координация.stack/Архивные.pulldown/Толщина воздуховодов.pushbutton/script.py`
- Batch fill + skip reasons: `WWBIM.extension/lib/Batch Operations/fill_section.py`, `WWBIM.extension/lib/Batch Operations/fill_floor.py`
- Type thickness retrieval patterns: `WWBIM.extension/lib/Batch Operations/fill_dimensions.py`

---

## Work Objectives

### Core Objective
Fill the following target parameters (all instance + Text) across the whole active document, strictly following (Category + optional `Группа модели`) rules.

### Confirmed Parameter Names (LookupParameter by name)
- `Группа модели`
- `ДГП_Площадь земельного участка`
- `ДГП_Площадь`
- `ДГП_Толщина`
- `ДГП_Высота`
- `ДГП_Ширина`
- `ДГП_Диаметр`
- `ДГП_Код помещения и зоны МССК`
- `ДГП_Наименование помещения и зоны МССК`
- `ADSK_Площадь с коэффициентом`
- `ADSK_Размер_Высота`
- `ADSK_Размер_Ширина`
- `ADSK_Размер_Диаметр`

### Rule Precedence (confirmed)
- If an element matches a group-specific rule (Category + `Группа модели`), apply ONLY that rule's assignments.
- Else apply category-default rules (no group condition) if defined.

### Formatting Rules (confirmed)
- Prefer project display string (`AsValueString()`) for numeric sources; if source is string, use `AsString()`.
- Strip unit suffix (e.g. `м²`, `мм`) and surrounding whitespace; handle NBSP.
- No extra rounding / no manual unit conversion.

### Write Policy (confirmed)
- Write only if the existing target value differs from expected (compare normalized strings); if equal -> skip.
- Missing target parameter -> report "параметр не найден" and skip (no parameter binding/creation).
- Missing/empty source -> skip and report a reason (no clearing of targets).

---

## Verification Strategy

### Automated Tests
- None (per request).

### QA Policy
- Script must produce a deterministic summary report (counts per rule + skip reasons + samples of failures) so it can be validated by log inspection.
- Evidence files (text) saved to `.sisyphus/evidence/` by the executor for each QA scenario.

---

## Execution Strategy

### Parallel Execution Waves

Wave 1 (Foundations + safety)
- Implement pure helpers (normalize/strip-units/compare)
- Add unit tests for helpers
- Update `bundle.yaml` text to match new feature intent

Wave 2 (Rule implementation, parallel by category)
- Implement collectors + rule dispatch + per-category rules

Wave 3 (Integration + reporting)
- Final logging/report formatting, performance pass, and evidence-driven QA

### Dependency Matrix (abbreviated)
- 1: —
- 2: —
- 3: —
- 4: 2, 3
- 5: 3, 4
- 6-12: 4, 5
- 13: 6-12
- F1: 13

---

## TODOs

- [ ] 1. Update `bundle.yaml` metadata to match DGP autofill

  **What to do**:
  - Change `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/bundle.yaml` title + tooltip to describe the new DGP parameter autofill behavior (not kubik sync).
  - Keep RU + EN blocks consistent (EN can be concise).

  **Must NOT do**:
  - Do not change script behavior here; metadata only.

  **Recommended Agent Profile**:
  - **Category**: `writing`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 1)
  - **Blocks**: none

  **References**:
  - `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/bundle.yaml` - currently mismatched tooltip (kubik sync).
  - `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/DGP_script.py` - target script for described behavior.

  **Acceptance Criteria**:
  - [ ] `bundle.yaml` RU title/tooltip describes Category+"Группа модели" rules and mentions "вся модель".

  **QA Scenarios**:
  ```
  Scenario: Metadata matches implemented feature
    Tool: Bash
    Steps:
      1. Open and review `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/bundle.yaml`
      2. Verify it no longer mentions kubik sync or "Отверстия" link
    Expected Result: Tooltip/title reflect DGP autofill feature
    Evidence: .sisyphus/evidence/task-1-bundle-yaml.txt
  ```

- [ ] 2. Add pure text normalization helpers for unit stripping + comparisons

  **What to do**:
  - Create `WWBIM.extension/lib/dgp_text.py` (pure-Python; no Revit imports) to normalize strings:
    - `normalize_group_value(s)` -> trimmed string (case preserved)
    - `strip_unit_suffix(s)` -> remove trailing unit tokens (`м²`, `m²`, `мм`, etc.), handle NBSP
    - `normalize_for_compare(s)` -> NBSP->space, collapse spaces, strip
    - `should_write(current, expected)` -> boolean

  **Must NOT do**:
  - No numeric parsing / no reformatting / no rounding.

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 1)
  - **Blocks**: Tasks 5-14 depend on these helpers

  **References**:
  - `WWBIM.extension/WW.BIM.tab/BIM.panel/Координация.stack/Архивные.pulldown/Толщина воздуховодов.pushbutton/script.py` - `AsValueString()`-driven display-string usage pattern.

  **Acceptance Criteria**:
  - [ ] Helper module exists and is importable with plain CPython (no Revit)
  - [ ] Handles inputs like `"1 234,56\u00A0м²"`, `"200 мм"`, `"  Газон  "`

  **QA Scenarios**:
  ```
  Scenario: Unit stripping is conservative
    Tool: Bash
    Steps:
      1. Run `python -c "import sys; sys.path.insert(0, r'WWBIM.extension\\lib'); import dgp_text; print(dgp_text.strip_unit_suffix('1 234,56\u00A0м²'))"`
    Expected Result: prints `1 234,56` (no trailing units)
    Evidence: .sisyphus/evidence/task-2-strip-units.txt
  ```

- [ ] 3. Add `unittest` coverage for helper functions
- [ ] 3. Define the DGP rule table in code (data-only)

  **What to do**:
  - Create `WWBIM.extension/lib/dgp_rules.py` (data-only; no Revit imports) that encodes all confirmed rules:
    - Category identifier (BuiltInCategory or a resolver key)
    - Optional group condition values (exact match after trim)
    - Target assignments (target name -> source spec)
  - Include rule precedence behavior explicitly (group-specific overrides category default).

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 1)
  - **Blocks**: Task 6-13

  **References**:
  - `.sisyphus/drafts/dgp-params.md` - source of truth for mapping rules.

  **Acceptance Criteria**:
  - [ ] Rules module contains all listed categories and group cases

  **QA Scenarios**:
  ```
  Scenario: Rules module imports in plain python
    Tool: Bash
    Steps:
      1. Run `python -c "import sys; sys.path.insert(0, r'WWBIM.extension\\lib'); import dgp_rules as r; print(len(r.RULES))"`
    Expected Result: prints a non-zero count
    Evidence: .sisyphus/evidence/task-3-rules-import.txt
  ```

- [ ] 4. Implement core Revit parameter IO helpers in `DGP_script.py`

  **What to do**:
  - In `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/DGP_script.py`, implement reusable helpers:
    - `get_param(elem, name)` and `get_type_param(elem, name)`
    - `get_value_as_display_text(param)`:
      - Prefer `param.AsValueString()` if non-empty; else fall back to StorageType-based reads.
      - Normalize/strip units using `WWBIM.extension/lib/dgp_text.py`.
    - `set_text_param_if_changed(elem, target_name, expected_text)`:
      - Enforce target StorageType String
      - Compare normalized strings; update only when changed
      - Return status reason (`updated`, `already_ok`, `parameter_not_found`, `wrong_storage_type`, `readonly`, `exception`)
  - Ensure per-element exceptions do not abort the run.

  **Must NOT do**:
  - Do not bind/create shared parameters.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2, after Tasks 2-4)
  - **Blocked By**: Tasks 2-3
  - **Blocks**: Tasks 6-13

  **References**:
  - `WWBIM.extension/WW.BIM.tab/Общее.panel/Кубики Синхронизация.pushbutton/Sync_script.py` - `get_param_value()` + `set_param_value()` patterns with StorageType.
  - `WWBIM.extension/lib/Batch Operations/fill_section.py` - `SetParameterValue()` status dict + skip reasons.

  **Acceptance Criteria**:
  - [ ] Helpers correctly distinguish missing target vs read-only vs wrong storage type
  - [ ] Helpers use unit stripping consistently

  **QA Scenarios**:
  ```
  Scenario: Helper module smoke test imports (no Revit)
    Tool: Bash
    Steps:
      1. Run `python -c "import sys; sys.path.insert(0, r'WWBIM.extension\\lib'); import dgp_text, dgp_rules"`
    Expected Result: imports succeed
    Evidence: .sisyphus/evidence/task-4-imports.txt
  ```

- [ ] 5. Implement rule dispatch engine (category + group precedence)

  **What to do**:
  - Implement:
    - group value read: `group = normalize_group_value(LookupParameter('Группа модели'))`
    - rule selection: group-specific match wins; else category-default rule
    - assignment execution: for each target in rule, resolve source (instance-first then type fallback for named params)
  - Make dispatch data-driven from `WWBIM.extension/lib/dgp_rules.py`.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2)
  - **Blocked By**: Tasks 3-4
  - **Blocks**: Tasks 7-13

  **References**:
  - `WWBIM.extension/lib/Batch Operations/fill_floor.py` - skip reasons accounting.

  **Acceptance Criteria**:
  - [ ] A single element never receives both group-specific and default assignments in same category
  - [ ] Group matching uses trim-only normalization, case-sensitive

  **QA Scenarios**:
  ```
  Scenario: Dispatch precedence smoke check
    Tool: Revit/pyRevit execution
    Steps:
      1. Run script on a small model with Floors both with and without `Группа модели` = `Газон`
      2. Verify elements matching the group-specific rule do not also receive default thickness fill
    Expected Result: Group-specific overrides default for the same category
    Evidence: .sisyphus/evidence/task-5-dispatch-smoke.txt
  ```

- [ ] 6. Implement TopographySurface rule (site area)

  **What to do**:
  - Collect `TopographySurface` elements and apply:
    - `ДГП_Площадь земельного участка` <- system/display `Площадь`

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2)
  - **Blocked By**: Tasks 5-6

  **Acceptance Criteria**:
  - [ ] Updates only when changed; missing targets reported

- [ ] 7. Implement Floors rules (Газон area vs default thickness)

  **What to do**:
- For `OST_Floors`:
    - If `Группа модели` == `Газон`: set `ДГП_Площадь` <- system/display `Площадь`
    - Else: set `ДГП_Толщина` <- type/display `Толщина`
  - Thickness source guidance:
    - Prefer type `LookupParameter('Толщина')` for display string
    - If missing, fall back to `BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM` pattern from `WWBIM.extension/lib/Batch Operations/fill_dimensions.py`

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2)
  - **Blocked By**: Tasks 5-6

- [ ] 8. Implement Walls rules (default thickness; group-specific фасад/окно)

  **What to do**:
- For `OST_Walls`:
    - If `Группа модели` == `Окно`: set `ДГП_Высота`/`ДГП_Ширина` from `ADSK_Размер_Высота`/`ADSK_Размер_Ширина`
    - If `Группа модели` == `Фасад`: set `ДГП_Толщина` <- type/display `Толщина`
    - Else: set `ДГП_Толщина` <- type/display `Толщина`
  - Thickness source guidance:
    - Prefer `WallType.Width` / type `LookupParameter('Толщина')` for display string
    - If missing, fall back to `BuiltInParameter.WALL_USER_WIDTH_PARAM` pattern from `WWBIM.extension/lib/Batch Operations/fill_dimensions.py`

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2)
  - **Blocked By**: Tasks 5-6

- [ ] 9. Implement Windows + Doors size rules

  **What to do**:
  - For `OST_Windows` and `OST_Doors`:
    - `ДГП_Высота` <- `ADSK_Размер_Высота`
    - `ДГП_Ширина` <- `ADSK_Размер_Ширина`
  - Source resolution: instance first, then type fallback.
  - Format as project display text; strip units.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2)
  - **Blocked By**: Tasks 5-6

- [ ] 10. Implement Structural Columns size rules

  **What to do**:
  - For `OST_StructuralColumns`:
    - `ДГП_Высота` <- `ADSK_Размер_Высота`
    - `ДГП_Ширина` <- `ADSK_Размер_Ширина`
    - `ДГП_Диаметр` <- `ADSK_Размер_Диаметр`
  - Source resolution: instance first, then type fallback.
  - If a particular source is missing (e.g. diameter for rectangular column), skip only that target and report reason.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2)
  - **Blocked By**: Tasks 5-6

- [ ] 11. Implement Rooms rule (area from `ADSK_Площадь с коэффициентом`)

  **What to do**:
  - For `OST_Rooms`:
    - `ДГП_Площадь` <- `ADSK_Площадь с коэффициентом`
  - If the source param is numeric: use `AsValueString()`; if string: use `AsString()`.
  - Strip unit suffix if present.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2)
  - **Blocked By**: Tasks 5-6

- [ ] 12. Implement Areas ("Зоны") rules for zoning plan

  **What to do**:
  - For `OST_Areas`:
    - If `Группа модели` == `СПП в ГНС`:
      - `ДГП_Код помещения и зоны МССК` = `9999`
      - `ДГП_Наименование помещения и зоны МССК` = `9999`
      - `ДГП_Площадь` <- system/display `Площадь`
    - If `Группа модели` == `Общая площадь`:
      - `ДГП_Код помещения и зоны МССК` = `П3 03`
      - `ДГП_Наименование помещения и зоны МССК` = `9999`
      - `ДГП_Площадь` <- system/display `Площадь`
    - Else: do nothing.
  - Enforce rule precedence: these are group-specific, so no other area rules apply.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES (Wave 2)
  - **Blocked By**: Tasks 5-6

- [ ] 13. Add deterministic reporting + evidence file output

  **What to do**:
  - Add a summary report at end of run:
    - totals per category rule: processed, updated, already_ok, missing_target, missing_source, readonly, wrong_storage_type, exceptions
    - sample element ids (bounded, e.g. first 20) for each failure reason
  - Write report to a file under `.sisyphus/evidence/` (timestamped) so QA can be non-interactive.
    - Resolve path relative to `__file__` (walk up to repo root) to find `.sisyphus/evidence/`.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: NO (Wave 3; integrates all rules)
  - **Blocked By**: Tasks 7-13

  **Acceptance Criteria**:
  - [ ] Running the script produces a report file with counts and skip reasons
  - [ ] Report path is printed to pyRevit output

  **QA Scenarios**:
  ```
  Scenario: Evidence report file is generated
    Tool: Revit/pyRevit execution
    Steps:
      1. Run the pushbutton "ДГП_Параметры" in a model containing at least one element for each implemented category
      2. Confirm script completes without unhandled exception
      3. Verify a new file exists under `.sisyphus/evidence/` matching `dgp-params-*.md`
    Expected Result: Report includes per-rule totals and skip reasons
    Evidence: .sisyphus/evidence/task-13-report-path.txt
  ```

---

## Final Verification Wave

- [ ] F1. Run script in Revit on a prepared fixture model

  **Tool**: Revit/pyRevit execution
  - Use a fixture RVT that contains at least:
    - 1 TopographySurface
    - 1 Floor with `Группа модели` = `Газон`
    - 1 Floor without group
    - 1 Wall with `Группа модели` = `Окно`
    - 1 Wall with `Группа модели` = `Фасад`
    - 1 Window, 1 Door, 1 Structural Column
    - 1 Room with `ADSK_Площадь с коэффициентом`
    - 2 Areas with `Группа модели` = `СПП в ГНС` and `Общая площадь`
  - Execute the pushbutton and verify:
    - no unhandled exceptions
    - report file exists under `.sisyphus/evidence/`
    - report shows non-zero processed counts for fixture categories
  - Save the report path and a short excerpt to `.sisyphus/evidence/final-revit-run.txt`

---

## Commit Strategy

- Commit 1: `docs(ui): update DGP button tooltip`
  - Files: `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/bundle.yaml`

- Commit 2: `feat(core): add DGP helper modules`
  - Files: `WWBIM.extension/lib/dgp_text.py`, `WWBIM.extension/lib/dgp_rules.py`

- Commit 3: `feat(dgp): fill DGP instance params by category and group`
  - Files: `WWBIM.extension/WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/DGP_script.py`

---

## Success Criteria

- Running the pushbutton in Revit completes and generates `.sisyphus/evidence/dgp-params-*.md`
- For each rule, at least one element in the fixture model is updated (when inputs exist) and skip reasons are reported otherwise
