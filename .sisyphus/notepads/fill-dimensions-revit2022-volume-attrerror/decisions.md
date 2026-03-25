# Decisions

## Task 2 Decision: Volume Source Replacement for Structural Elements

### Problem
- `BuiltInParameter.STRUCTURAL_VOLUME` does not exist in Revit 2022 API
- Caused `AttributeError` at import time (module-level constant evaluation)

### Solution
- Replaced `BuiltInParameter.STRUCTURAL_VOLUME` with `BuiltInParameter.HOST_VOLUME_COMPUTED` for `OST_StructuralColumns` and `OST_StructuralFraming` (line 46)
- `HOST_VOLUME_COMPUTED` exists in Revit 2022 and is a valid parameter for volume
- This change resolves the import-time `AttributeError`

### Rationale
- Task 3 will implement proper import-safe enum resolution
- For now, use `HOST_VOLUME_COMPUTED` as a minimal fix to allow import
- Keeps the same categories (`OST_StructuralColumns`, `OST_StructuralFraming`)

## Task 3 Decision: Import-safe BuiltInParameter Resolution

- Converted `DIMENSION_PARAMETERS[*]["SOURCES"][*]["BIP"]` values from direct `BuiltInParameter.*` enum access to string names, so module import no longer evaluates optional enum members.
- Added `_resolve_bip(name)` using `Enum.IsDefined(BuiltInParameter, name)` + `Enum.Parse(...)` and applied resolution inside `_BuildSourceMapForParam`.
- `_BuildSourceMapForParam` now caches resolved enum values by mutating `source["BIP"]` when resolution succeeds; unresolved names are skipped gracefully without raising.
