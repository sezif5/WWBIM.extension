# Fix fill_dimensions.py Revit 2022 AttributeError (Structural Volume)

## TL;DR
> `WWBIM.extension/lib/Batch Operations/fill_dimensions.py` падает при старте в Revit 2022 из-за обращения к несуществующему члену `BuiltInParameter.*` для объема конструкций. Нужно убрать обращение к отсутствующим enum-атрибутам на этапе импорта и заменить источник объема для несущих категорий на реально существующий системный параметр (в первую очередь `HOST_VOLUME_COMPUTED`), сохранив текущую логику заполнения.

**Deliverables**:
- Скрипт `WWBIM.extension/lib/Batch Operations/fill_dimensions.py` импортируется и выполняется в batch runner на Revit 2022 без AttributeError.
- `ADSK_Размер_Объём` продолжает заполняться для несущих колонн/балок из системного параметра объема (без геометрического пересчета по умолчанию).

**Estimated Effort**: Short
**Parallel Execution**: YES (2 waves)
**Critical Path**: Confirm executed script copy -> make BIP resolution import-safe -> validate in Revit 2022 batch runner

---

## Context

### Original Request
Ошибка при запуске `D:\GitHub\WWBIM.extension\WWBIM.extension\lib\Batch Operations\fill_dimensions.py`: `type object has no attribute 'Structutral_VOLUME'`; скрипт не выполняется. Revit 2022, запуск через batch runner. Другие скрипты с тем же `add_shared_parameter.py` работают.

### Key Findings
- `WWBIM.extension/lib/Batch Operations/fill_dimensions.py` содержит ссылку на `BuiltInParameter.STRUCTURAL_VOLUME` в константе `DIMENSION_PARAMETERS` (импорт-тайм) — любое отсутствие атрибута ломает импорт до `Execute(...)`.
- Для несущих колонн пользователь ожидает использовать системный параметр объема (в UI он есть).

### Metis Guardrails (incorporated)
- Избежать падений на импорте: никакого прямого доступа к потенциально отсутствующим `BuiltInParameter.*` при загрузке модуля.
- Не менять существующую бизнес-логику заполнения (конвертации, приоритет источников) кроме части, необходимой для совместимости.
- Не включать расчет объема по геометрии как дефолт (риск деградации производительности batch).

---

## Work Objectives

### Core Objective
Сделать `fill_dimensions.py` совместимым с Revit 2022: исключить AttributeError при старте и корректно брать объем для несущих категорий из доступного системного параметра.

### Definition of Done
- Скрипт загружается и выполняется в batch runner (Revit 2022) без исключений.
- На тестовой модели есть хотя бы 1 элемент `OST_StructuralColumns` и 1 элемент `OST_StructuralFraming`; после выполнения у них заполнен `ADSK_Размер_Объём` (не пусто, значение > 0 для заведомо ненулевых элементов).

---

## Verification Strategy

### Test Decision
- **Infrastructure exists**: Revit batch runner (pyRevit) сценарии
- **Automated tests**: None (Revit API зависимость)
- **Primary verification**: agent-executed run inside Revit 2022 + evidence capture (batch runner logs / screenshots)

### QA Policy (evidence)
Сохранять результаты выполнения в `.sisyphus/evidence/`:
- `task-1-import-ok.txt` (лог загрузки модуля)
- `task-4-run-ok.txt` (лог результата `Execute` + счетчики)
- `task-4-volume-filled.txt` (вывод проверки 2 элементов: category/id + значение параметра)

---

## Execution Strategy

Wave 1 (can start immediately)
- Task 1: Confirm the *exact* script copy and failing line in batch runner
- Task 2: Replace invalid structural volume BIP with a valid source for Revit 2022
- Task 3: Make all BIP references import-safe (runtime resolution)
- Task 4: Revit 2022 batch runner QA on a minimal model
- Task 5: Document the compatibility rule (so it doesn't regress)

Wave 2 (after Wave 1)
- Task 6: Optional performance check on a representative model set (timings)

---

## TODOs

- [x] 1. Confirm executed script path and failing attribute

  **What to do**:
  - Reproduce failure in the same runner used by user.
  - Capture stack trace line showing `File ...fill_dimensions.py, line ...` and the exact missing attribute name.
  - Verify which copy of `fill_dimensions.py` is loaded (repo path vs deployed extension path).

  **References**:
  - `WWBIM.extension/lib/Batch Operations/fill_dimensions.py` - has structural volume source mapping.
  - `WWBIM.extension/WW.BIM.tab/BIM.panel/Пакетные операции.pushbutton/script.py` - batch runner entrypoint wiring.

  **Acceptance Criteria / QA Scenarios**:
  ```
  Scenario: Capture exact failing attribute and path
    Tool: Revit 2022 batch runner log
    Steps:
      1. Run batch operation that triggers fill_dimensions
      2. Save the full exception output
    Expected Result: Log contains exact file path + line number + missing attribute name
    Evidence: .sisyphus/evidence/task-1-import-ok.txt
  ```

- [x] 2. Fix structural volume source for Revit 2022 (use system volume)

  **What to do**:
  - Remove/replace use of `BuiltInParameter.STRUCTURAL_VOLUME` for `OST_StructuralColumns` and `OST_StructuralFraming`.
  - Prefer using the system computed volume parameter available in Revit 2022 (default: `BuiltInParameter.HOST_VOLUME_COMPUTED`).
  - Keep the existing “first non-empty source wins” behavior.

  **Must NOT do**:
  - Do not switch to geometry volume calculation by default.
  - Do not change unit conversions in `FormatValue`.

  **References**:
  - `WWBIM.extension/lib/Batch Operations/fill_dimensions.py` - `DIMENSION_PARAMETERS` volume section.

  **Acceptance Criteria / QA Scenarios**:
  ```
  Scenario: No missing BIP for structural volume
    Tool: Revit 2022 batch runner
    Steps:
      1. Import/load the script module
    Expected Result: No AttributeError about structural volume
    Evidence: .sisyphus/evidence/task-1-import-ok.txt
  ```

- [x] 3. Make BuiltInParameter resolution import-safe (cross-version guard)

  **What to do**:
  - Ensure `fill_dimensions.py` does not access `BuiltInParameter.SOMETHING` at import time for any potentially missing members.
  - Store BIP identifiers as strings (e.g. `"HOST_VOLUME_COMPUTED"`) or resolve via safe `getattr`/`Enum.Parse` at runtime.
  - Ensure `_BuildSourceMapForParam` skips missing/None BIPs.

  **Must NOT do**:
  - Do not silently swallow errors without recording a reason (at least optionally return counters/log message).

  **References**:
  - `WWBIM.extension/lib/Batch Operations/fill_dimensions.py` - `_BuildSourceMapForParam`, `GetBuiltInParamValue`.

  **Acceptance Criteria / QA Scenarios**:
  ```
  Scenario: Module imports even if a BIP name is not available
    Tool: Revit 2022 batch runner
    Steps:
      1. Load module
      2. Run Execute(doc)
    Expected Result: Execute returns dict with success=True; missing BIPs are skipped (no crash)
    Evidence: .sisyphus/evidence/task-4-run-ok.txt
  ```

- [x] 4. Revit 2022 batch runner QA: verify volume filled for structural elements

  **What to do**:
  - Run `Execute(doc)` on a minimal model containing at least 1 structural column and 1 structural framing element.
  - After run, read `ADSK_Размер_Объём` value from both elements and log it.

  **References**:
  - `WWBIM.extension/lib/Batch Operations/fill_dimensions.py:Execute` - return contract and transaction behavior.

  **Acceptance Criteria / QA Scenarios**:
  ```
  Scenario: Structural column and framing get ADSK_Размер_Объём
    Tool: Revit 2022 (batch runner script / RevitPythonShell)
    Steps:
      1. Pick one element from OST_StructuralColumns and OST_StructuralFraming
      2. Run fill_dimensions.Execute(doc)
      3. Read LookupParameter("ADSK_Размер_Объём").AsString() for both
    Expected Result: Both values are non-empty; for known-nonzero elements value parses as float > 0
    Evidence: .sisyphus/evidence/task-4-volume-filled.txt
  ```

- [x] 5. Add a short note to Batch Operations documentation about Revit-version-safe BIPs

  **What to do**:
  - Update `WWBIM.extension/lib/Batch Operations/README.md` with a warning: avoid direct `BuiltInParameter.X` in module-level constants; resolve at runtime.
  - Mention which BIP is used for structural volume in Revit 2022.

  **Acceptance Criteria / QA Scenarios**:
  ```
  Scenario: Documentation updated
    Tool: file diff review
    Steps:
      1. Confirm README contains the note and points to fill_dimensions behavior
    Expected Result: Guidance exists to prevent regressions
    Evidence: .sisyphus/evidence/task-5-doc-note.txt
  ```

- [x] 6. Optional: Define minimal timing capture procedure in batch runner
2. Run baseline on a known-good model.
2. Execute `fill_dimensions.py` on the test model.
2. Record start time and element processed.
3. After completion, add timing data to evidence log

4. Save evidence file

---

**Timing Evidence Template:**

```
=== Performance Test ===
Date: 
Model: 
Start Time: 
End Time: 
Duration: 
Elements Processed: 
Elements Updated: 
Elements Skipped: 

Per-Parameter Breakdown:
| Parameter | Updated | Skipped | No Value |
|---|-------|-------|--------|---------|
| ADSK_Размер_Объём |    Y    |    Y    |    N    |    Y        |
|    Y    |    Y        |    N        |    Y        |    N        |    Y        |    N        |    N       |
|    N        |    N         |    N        |
|    Y       |    Y        |    N        |    N       |
|    N       |    N        |    N        |
|    Y       |    Y        |    N        |    Y       |
|    N       |    N        |    N        |    N        |
|    N       |    N        |    N        |    N        |
|    N       |    N         |    N        |    N        |
|    N       |    N         |    N        |    N        |
+------------------------------+
Total: 
Duration: 

Per-Parameter Stats:
| Parameter | Updated | Skipped | No Value |
|---|-------|-------|--------| ---------|
| ADSK_Размер_Объём |    Y    |    Y        |    N        |    Y        |
|    N        |    Y        |    N        |    N       |    N        |    N        |    N       |    N        |
|    N       |    N        |    N        |    N        |
|    N       |    N        |    N        |    N        |
|    N       |    N         |    N        |    N        |
+------------------------------+
Total: 
Duration: 
```

*Generated: 2026-03-23*
*Plan reference: .sisyphus/plans/fill-dimensions-revit2022-volume-attrerror.md*

  **What to do**:
  - Run the batch operation on a representative model set.
  - Compare total runtime before/after (or at least record per-model timings now).

  **Acceptance Criteria / QA Scenarios**:
  ```
  Scenario: No major batch slowdown
    Tool: batch runner logs
    Steps:
      1. Run on N models
      2. Record timings
    Expected Result: Runtime delta within agreed tolerance (default target: <= +10%)
    Evidence: .sisyphus/evidence/task-6-timings.txt
  ```

---

## Final Verification Wave
- Re-run Task 4 QA scenario on a second model (different file) to confirm it isn't model-specific.

## Commit Strategy
- One commit: `fix(batch): make fill_dimensions Revit 2022-safe for volume BIP`

## Success Criteria
- No startup AttributeError in Revit 2022 batch runner.
- Structural categories keep filling `ADSK_Размер_Объём` from system computed volume.
