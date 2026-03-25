# Draft: DGP parameters autofill

## Requirements (unconfirmed)
- Need a Revit/pyRevit script in `WW.BIM.tab/АГР.panel/ДГП_Параметры.pushbutton/DGP_script.py`.
- Script fills *instance* shared parameters ("общие параметры экземпляра").
- The filled parameter is always a Text parameter.
- The source parameter (copied-from) can be numeric; value must be converted to text.
- Filling logic depends on element Category and the content of a parameter like "Группа модели".

## Requirements (confirmed)
- Button goal: **fill instance shared parameters based on Category + "Группа модели"** (NOT the "Sync Kubiks" behavior currently described in `bundle.yaml`).

## Process Constraint (confirmed)
- User requests: follow the plan but **no automated testing tasks**; focus on implementing the script.

## Mapping Rules (confirmed so far)
- Category: `Топография`
  - Target (instance shared, Text): `ДГП_Площадь земельного участка`
  - Source (system, numeric): `Площадь`
  - Behavior: copy source value into target, converting to text.

- Category: `Перекрытия` with "Группа модели" == "Газон" (group-model condition applies)
  - Target (instance shared, Text): `ДГП_Площадь`
  - Source (system, numeric): `Площадь`
  - Behavior: copy project display value (`AsValueString()`), strip unit suffix, no extra rounding.

- Category: `Перекрытия` (general)
  - Target (instance shared, Text): `ДГП_Толщина`
  - Source (system, type numeric): `Толщина`

- Category: `Стены` (general)
  - Target (instance shared, Text): `ДГП_Толщина`
  - Source (system, type numeric): `Толщина`

- Category: `Стены` with "Группа модели" == "Фасад"
  - Target (instance shared, Text): `ДГП_Толщина`
  - Source (system, type numeric): `Толщина`

- Category: `Окна`
  - Target (instance shared, Text): `ДГП_Высота`
  - Target (instance shared, Text): `ДГП_Ширина`
  - Source: `ADSK_Размер_Высота`, `ADSK_Размер_Ширина`

- Category: `Двери`
  - Target (instance shared, Text): `ДГП_Высота`
  - Target (instance shared, Text): `ДГП_Ширина`
  - Source: `ADSK_Размер_Высота`, `ADSK_Размер_Ширина`

- Category: `Стены` with "Группа модели" == "Окно"
  - Target (instance shared, Text): `ДГП_Высота`
  - Target (instance shared, Text): `ДГП_Ширина`
  - Source: `ADSK_Размер_Высота`, `ADSK_Размер_Ширина`

- Category: `Несущие колонны` (Structural Columns)
  - Target (instance shared, Text): `ДГП_Высота` from `ADSK_Размер_Высота`
  - Target (instance shared, Text): `ДГП_Ширина` from `ADSK_Размер_Ширина`
  - Target (instance shared, Text): `ДГП_Диаметр` from `ADSK_Размер_Диаметр`

- Category: `Помещения`
  - Target (instance shared, Text): `ДГП_Площадь`
  - Source: a parameter that already contains "area with coefficient"; no extra multiplication required.
  - Behavior: copy source value into target as text.

## Parameter Names (confirmed)
- Rooms source parameter: `ADSK_Площадь с коэффициентом`
 - Dimensions sources: `ADSK_Размер_Высота`, `ADSK_Размер_Ширина`, `ADSK_Размер_Диаметр`
 - Dimensions targets: `ДГП_Высота`, `ДГП_Ширина`, `ДГП_Диаметр`, `ДГП_Толщина`

## Formatting Decisions (confirmed)
- Rooms: format like other areas: numeric -> `AsValueString()` (project display), strip unit suffix; string -> use value and strip unit suffix if present.

- Category: `Зона`
  - Applies only when parameter "Группа модели" has specific values:
    - If "Группа модели" == "СПП в ГНС":
      - `ДГП_Код помещения и зоны МССК` = "9999"
      - `ДГП_Наименование помещения и зоны МССК` = "9999"
      - `ДГП_Площадь` = system `Площадь` (project display via `AsValueString()`, unit suffix stripped)
    - If "Группа модели" == "Общая площадь":
      - Same as above, but `ДГП_Код помещения и зоны МССК` = "П3 03" (instead of 9999)
      - `ДГП_Наименование помещения и зоны МССК` = "9999" (unchanged)

## Category Clarification (confirmed)
- "Зона" (for zoning plan) means Areas: `OST_Areas`.

## Category Clarification (confirmed)
- "Топография" means `TopographySurface` (not Toposolid).

## Formatting Decisions (partially confirmed)
- User wants: "без округления".
- User clarified: copy the system parameter value **without manual conversion** and **without "м2"** in the target text.

## Formatting Decisions (confirmed)
- For `Площадь`: write the **project display value** using `AsValueString()` and strip the unit suffix (e.g. remove trailing "м²" / spaces). Do not apply additional rounding.
- For dimensions (`ДГП_Толщина/Высота/Ширина/Диаметр`): use the same approach: prefer `AsValueString()` / existing string, strip unit suffix, no extra rounding.

## Normalization (default)
- When stripping units / comparing values: handle NBSP and multiple spaces; compare normalized `strip()`ed strings.

## Group Model Usage (partially confirmed)
- "Группа модели" is currently used at least for: Floors (Перекрытия) when value == "Газон".
- "Группа модели" is used for: Zone elements when value == "СПП в ГНС" or "Общая площадь".

## Group Model Matching (confirmed)
- Matching rule for "Группа модели" conditions: strict string equality.

## Group Model Normalization (confirmed)
- When comparing "Группа модели": trim whitespace at both ends before equality check; case sensitivity preserved.

## Rule Precedence (confirmed intent)
- User clarified: do NOT "always fill everything"; apply rules within (Category + optional "Группа модели") so that specific group rules can exclude other fills.

## Rule Precedence (proposal)
- Within a category: if an element matches a group-specific rule (exact match after trim), apply **only that rule's assignments**.
- If no group-specific rule matches: apply the category default rule(s) that have no group condition.

## Processing Scope (confirmed)
- Apply to: **entire active model/document** (not just selection).

## Write Behavior (confirmed)
- Target `ДГП_Площадь земельного участка`: skip write **only if** current value already equals the computed/expected value; otherwise overwrite.

## Source Missing Behavior (default)
- If a source parameter is missing/empty/unreadable for a rule: do not change the target; log a skip reason (e.g. "source_not_found"/"source_empty").

## Source Resolution (default)
- For named source parameters (e.g. `ADSK_Размер_Высота`): lookup order = instance first (`element.LookupParameter`), then type (`element.Document.GetElement(element.GetTypeId())`).

## Missing Parameter Handling (confirmed)
- If target param `ДГП_Площадь земельного участка` is not found on an element: log/report "параметр не найден" and skip; do not auto-bind/create.

## Missing Parameter Handling (assumption)
- Default for other target params: same behavior (log + skip; no auto-bind). If this is wrong, we need to confirm.

## Remaining Ambiguity
- For Rooms (Помещения): exact source parameter name (you mentioned both `ADSK_Площадь` and "ADSK_Площадь с коэффициентом") and its StorageType (String vs Double).
- For Rooms (Помещения): should we apply the same formatting rule as system `Площадь` (use `AsValueString()` and strip unit suffix), or copy raw string as-is?
- For Floors "Газон": match rule is equals vs contains vs startswith; case sensitivity.
- Category "Зона": confirm exact Revit element type/category used for "План зонирования" (likely Areas), so element collection is correct.
- Zone "Общая площадь": confirm value for `ДГП_Наименование помещения и зоны МССК` (keep "9999" vs other).

## Open Questions
- For numeric source → text target, should we copy the *display string* (`AsValueString`, e.g. "123,45 м²") or a *plain number* without units? If plain: rounding rules + decimal separator.
- Processing scope default: all elements in active document vs current selection vs active view.
- If `ДГП_Площадь земельного участка` is missing/unbound, should script auto-bind it using `AddSharedParameterToDoc()` or just report and skip?

## Repo Findings
- `D:\GitHub\WWBIM.extension\WWBIM.extension\WW.BIM.tab\АГР.panel\ДГП_Параметры.pushbutton\DGP_script.py` is currently empty.
- `D:\GitHub\WWBIM.extension\WWBIM.extension\WW.BIM.tab\АГР.panel\ДГП_Параметры.pushbutton\bundle.yaml` tooltip describes a *different* behavior: sync opening-task ("кубики") parameters from a linked model with "Отверстия" in its name, matching by parameter "№ кубика", and updating Status / Approval flags / Comments.

- Existing patterns to reuse:
  - `D:\GitHub\WWBIM.extension\WWBIM.extension\WW.BIM.tab\Общее.panel\Кубики Синхронизация.pushbutton\Sync_script.py` has robust helpers for reading/writing parameters across StorageTypes:
    - `get_param_value()` handles `String/Integer/Double/ElementId`
    - `get_param_value_string()` prefers `AsValueString()` then falls back
    - `set_param_value()` sets based on target StorageType and returns `(success, error)`
  - `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_section.py` shows batch-filling a shared *instance* text parameter using a mapping file + `AddSharedParameterToDoc()`:
    - `EnsureParameterExists()` binds shared parameter to `MODEL_CATEGORIES`
    - `GetParameterValue()`/`SetParameterValue()` enforce target StorageType==String and return structured status/reasons
    - Reads mapping from `Objects/*.txt`
  - `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_floor.py` repeats the same Set/Get pattern and demonstrates skip-reasons accounting for large runs.
  - `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\add_shared_parameter.py` provides `AddSharedParameterToDoc()` to bind shared params by name/GUID with diagnostics.
  - `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\model_categories.py` provides `MODEL_CATEGORIES` for multi-category filtering.

## Additional Repo Findings
- No existing python references to `ДГП_` parameters (new to implement).
- `ADSK_Размер_Высота` appears in `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_dimensions.py` (as a target filled parameter) and in `...\Универсальное семейство.pushbutton\script.py`.
- No existing python references to parameter name `Группа модели`; logic will be new.
- Useful formatting pattern: `D:\GitHub\WWBIM.extension\WWBIM.extension\WW.BIM.tab\BIM.panel\Координация.stack\Архивные.pulldown\Толщина воздуховодов.pushbutton\script.py` uses `AsValueString()` for getting display strings.

## Test Infrastructure
- No obvious existing test setup found (no `pytest.ini`, `pyproject.toml`, `.github/workflows`, or `*test*.py` discovered via glob).

## Technical Decisions (pending)
- Revit element selection scope: [DECISION NEEDED]
- Mapping rules (Category + Group value → which target params to set from which sources): [DECISION NEEDED]
- Numeric-to-text formatting rules (decimal separator, rounding, units): [DECISION NEEDED]
- Write behavior when target already has value: [DECISION NEEDED]

## Open Questions
- Which categories are in scope, and how exactly should "Группа модели" be interpreted (equals/contains/regex)?
- Which exact parameter names: group-param, source param(s), target param(s) (shared vs built-in)?
- Apply to all instances in model / active view / current selection?
- Skip element types? Handle linked models?

## Scope Boundaries
- INCLUDE: deterministic parameter mapping + safe conversion to text + logging/reporting.
- EXCLUDE (unless requested): creation/binding of new shared parameters; editing families; processing linked docs.
