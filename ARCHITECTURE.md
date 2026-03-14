# Architecture

## Overview
- pyRevit extension for Autodesk Revit that adds WW.BIM tab tools, background exporters, and batch parameter utilities.
- Primary runtime is IronPython via pyRevit, with deep Revit API usage and optional .NET UserControl dockable pane.

## Tech Stack
- Language: Python (IronPython/pyRevit runtime).
- Platform APIs: Autodesk Revit API, .NET (clr), pyRevit framework.
- UI: pyRevit buttons/panels; optional WPF UserControl via DLL.

## Directory Structure
- `WWBIM.extension/` - pyRevit extension root.
- `WWBIM.extension/startup.py` - loads FamilyManager DLL and registers dockable pane.
- `WWBIM.extension/nwc_export_timer.py` - daily timer that fires a pyRevit ExternalEvent.
- `WWBIM.extension/lib/` - shared helpers (background open/close, export, params).
- `WWBIM.extension/lib/Batch Operations/` - batch parameter fill scripts + docs.
- `WWBIM.extension/WW.BIM.tab/` - pyRevit UI layout, panels, and button entry scripts.

## Main Entry Points
- `WWBIM.extension/startup.py` - pyRevit startup hook; loads `FamilyManager.dll` and registers a dockable pane.
- `WWBIM.extension/nwc_export_timer.py` - creates a .NET timer and sends `daily_nwc_export` ExternalEvent.
- UI scripts under `WWBIM.extension/WW.BIM.tab/**.pushbutton/` - per-tool command entry points.

## Core Components
- Background open/close
  - `WWBIM.extension/lib/openbg.py` - opens RVT files with workset filtering and dialog/warning suppression.
  - `WWBIM.extension/lib/closebg.py` - closes/synchronizes/saves documents with safe fallback behavior.
- Export pipeline
  - `WWBIM.extension/lib/nwc_export_utils.py` - shared logic for Navisworks export (view setup, export, metrics).
  - `WWBIM.extension/lib/export_single_rvt_to_nwc.py` - wraps single-file export flow.
  - `WWBIM.extension/WW.BIM.tab/BIM.panel/Экспорт.stack/` - UI scripts for manual/auto export.
- Batch operations
  - `WWBIM.extension/lib/Batch Operations/*.py` - parameter fill/copy utilities.
  - `WWBIM.extension/lib/Batch Operations/README.md` - behavior documentation.
- Parameters and category helpers
  - `WWBIM.extension/lib/add_shared_parameter.py` - shared parameter binding utilities.
  - `WWBIM.extension/lib/model_categories.py` - category lists used by batch tools.

## Data Flow
- UI scripts collect user input or read object lists, then:
  - Open models with `openbg.open_in_background`.
  - Prepare a dedicated 3D view for export (hide annotations/links/imports).
  - Export NWC/RVT via Revit API methods.
  - Close documents via `closebg.close_with_policy`.
- Batch operations:
  - Read model data → bind shared parameters → compute values → write parameters.
  - Return a summary dict for logging/aggregation.

## External Integrations
- Autodesk Revit API (Open/Export/Save/Synchronize, transactions, view creation).
- pyRevit APIs (script output/logging, events, forms).
- .NET assembly loading (`FamilyManager.dll`) in `WWBIM.extension/startup.py`.
- Filesystem and network paths:
  - Export object lists and output folders via `Y:\BIM\Scripts\Objects\...`.
  - Local cache under `%LOCALAPPDATA%\pyRevit\FamilyManager\cache`.

## Configuration
- Startup/DLL:
  - `WWBIM.extension/startup.py` - `DLL_PATHS`, `CLASS_FULLNAME`, `PANE_GUID`, `PANE_TITLE`.
- Timer:
  - `WWBIM.extension/nwc_export_timer.py` - `EXPORT_TIME`, `CHECK_INTERVAL_MINUTES`, `EXTERNAL_EVENT_NAME`.
- Export scripts:
  - `WWBIM.extension/WW.BIM.tab/BIM.panel/Экспорт.stack/АвтоЭкспортRVT.pushbutton/auto_rvt_export_script.py` - `OBJECTS_FILE`, `OBJECTS_BASE_DIR`, logging paths.
  - `WWBIM.extension/WW.BIM.tab/BIM.panel/Экспорт.stack/АвтоЭкспортNWC.pushbutton/auto_navis_export_script.py` - `OBJECTS_FILE`, `OBJECTS_BASE_DIR`, logging paths.

## Build & Deploy
- Deployment model: pyRevit extension folder, loaded by pyRevit at Revit startup.
- No build step for Python scripts; optional DLLs must exist at `DLL_PATHS`.
- No CI/lint configs detected in repository root.

## Tests
- No test suites detected; only a cached bytecode file under `__pycache__/`.
