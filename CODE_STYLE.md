# Code Style

## Naming Conventions
- Files
  - Library modules use lower_snake_case: `WWBIM.extension/lib/openbg.py`, `WWBIM.extension/lib/nwc_export_utils.py`.
  - UI entry scripts often end in `_script.py` and use descriptive names: `WWBIM.extension/WW.BIM.tab/Оформление.panel/ВыравнитьВиды.pushbutton/AlignViews_script.py`.
  - Batch operations grouped under `WWBIM.extension/lib/Batch Operations/`.
- Classes
  - PascalCase for types and Revit API helpers: `WWBIM.extension/lib/openbg.py` (`SuppressWarningsPreprocessor`).
  - Private/internal classes may use leading underscore: `WWBIM.extension/startup.py` (`_Provider`).
- Functions
  - snake_case for helpers: `WWBIM.extension/lib/nwc_export_utils.py` (`to_model_path`, `export_view_to_nwc`).
  - PascalCase for API-like utilities in some modules: `WWBIM.extension/lib/add_shared_parameter.py` (`FindSharedParameterElementByGuid`).
- Constants
  - UPPER_SNAKE_CASE at module level: `WWBIM.extension/lib/add_shared_parameter.py` (`CONFIG`).
  - Script-level settings in export scripts: `WWBIM.extension/WW.BIM.tab/BIM.panel/Экспорт.stack/АвтоЭкспортNWC.pushbutton/auto_navis_export_script.py` (`OBJECTS_FILE`, `LOG_DIR`).

## File Organization
- Shared reusable logic lives in `WWBIM.extension/lib/`.
- Command entry scripts live under `WWBIM.extension/WW.BIM.tab/` with `.panel/.stack/.pushbutton` folders.
- Batch tools are grouped and documented in `WWBIM.extension/lib/Batch Operations/README.md`.

## Import Style
- Standard library imports first (`os`, `sys`, `datetime`, `codecs`).
- pyRevit imports next (`pyrevit.script`, `pyrevit.coreutils`, `pyrevit.events`, `pyrevit.forms`).
- Revit API imports grouped from `Autodesk.Revit.DB` and `Autodesk.Revit.UI`.
- Local libs imported by name after `sys.path` adjustments in UI scripts:
  - Example: `WWBIM.extension/WW.BIM.tab/BIM.panel/Экспорт.stack/АвтоЭкспортRVT.pushbutton/auto_rvt_export_script.py` adds `lib_path` and imports `openbg`, `closebg`.

## Code Patterns
- Broad Revit API safety
  - Extensive `try/except` blocks around Revit API calls to keep tools resilient.
  - Example: `WWBIM.extension/lib/openbg.py` suppresses dialog failures and retries open with alternate workset configs.
- Workset filtering
  - Predicate excludes worksets starting with `00_` and containing `Link`/`Связь`.
  - Example: `WWBIM.extension/lib/nwc_export_utils.py` (`workset_filter`).
- Export pipeline
  - Prepare 3D view, hide annotations/imports/links, export, then close.
  - Example: `WWBIM.extension/lib/nwc_export_utils.py` (`find_or_create_navis_view`, `export_view_to_nwc`).
- Batch operation contracts
  - Functions return a result dict with `success`, `message`, `parameters`, `fill` keys.
  - Documented in `WWBIM.extension/lib/Batch Operations/README.md`.

## Error Handling
- Prefer defensive `try/except` with graceful fallback and logging.
- Export flows continue on per-model failure and record counts in logs.
- Examples:
  - `WWBIM.extension/lib/closebg.py` falls back to SaveAs temp-file on errors.
  - `WWBIM.extension/lib/openbg.py` suppresses failure handler exceptions.

## Logging
- pyRevit logger for startup and shared libs: `WWBIM.extension/startup.py` (`script.get_logger()`).
- Script logs via `print()` and file logs using `codecs.open` for UTF-8.
- Dialog suppression logs collected into lists and summarized.

## Testing
- No tests or testing framework detected in repo.

## Do's and Don'ts
- Do follow the existing workset filter logic in new background-open code.
- Do use `openbg.open_in_background` and `closebg.close_with_policy` for background document work.
- Do wrap Revit API calls with `try/except` to avoid blocking user workflows.
- Don't assume all Revit BuiltInCategory values exist; use `Enum.IsDefined` like `WWBIM.extension/lib/openbg.py`.
- Don't introduce new config files without wiring them into scripts that read from filesystem paths.
