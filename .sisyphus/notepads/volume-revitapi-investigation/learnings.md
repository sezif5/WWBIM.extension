# Volume Retrieval Methods in Revit API - Learnings

## 6 Credible References Found (2022-2026)

### 1. GetMaterialVolume() - Performance-Optimized Approach
**Source**: Autodesk/revit-ifc - CategoryUtil.cs#L608 (2026)
- Uses internal material quantity calculation - NO geometry extraction
- Applies to Category.HasMaterialQuantities == true elements
- Works for: walls, roofs, floors, ceilings, 3D families (columns, MEP equipment)
- Returns volume in cubic feet (internal units)
- Code: `element.GetMaterialVolume(materialId)`
- Best for batch runner performance

### 2. BuiltInParameter.STEEL_ELEM_PROFILE_VOLUME
**Source**: Revit API Docs 2022 + CodeCavePro decompiled
- Language-independent enum approach (safer than name lookup)
- Enum value: -1155148 (0xFFEE5FB4)
- Specific to steel structural framing profiles
- Revit 2022+ compatible
- Code: `element.get_Parameter(BuiltInParameter.STEEL_ELEM_PROFILE_VOLUME)?.AsDouble()`

### 3. LookupParameter("Volume") - Name-Based
**Source**: Autodesk Forum 2022 discussion
- Returns first match if multiple params have same name
- NOT language-portable (built-in names translated)
- Jeremy Tammik: avoid unless you can guarantee unique names
- Safer alternatives: GetParameters("Volume") or get_Parameter(BuiltInParameter)
- Code: `element.LookupParameter("Volume")?.AsDouble()`

### 4. Geometry-Based Solid.Volume
**Source**: Autodesk Forum - Compute Intersection Volume (2022)
- Returns signed volume from actual geometry
- Revit computes analytically when possible, tessellation fallback for curved surfaces
- May be slightly under/overestimated with curved surfaces
- Performance: EXPENSIVE - geometry extraction is heavy
- Code: `element.get_Geometry(opt)` → iterate to find Solid → `solid.Volume`

### 5. UI Volume vs Computed Volume - Detail Level Issue
**Source**: Autodesk Support 2023 + BIMNature Blog 2018
- Material:Volume = based on Fine detail level
- Volume = based on Coarse detail (beams fall back to Medium)
- **Medium detail objects NEVER counted** in either calculation
- Example: 1.0m³ (Coarse) + 0.5m³ (Medium) + 0.25m³ (Fine) = 1.75m³ actual
  - Schedule shows different values because 0.5m³ is ignored
- Use GetMaterialVolume() for consistency with Fine detail

### 6. Material Takeoff Weight Calculation
**Source**: Revit Waterman Blog 2026
- Formula: `Material:Volume / 1 * 7.85` (7.85 g/cm³ steel density)
- Divide by "1" to neutralize units (avoids "Inconsistent Units" error)
- OOTB families use medium detail → missing root radii = incorrect weights
- Solution: custom families with full profile in both medium AND fine

## Methods Used

- GetMaterialVolume(): ✅ Performance-optimized, ✅ Accurate (Fine detail)
- BuiltInParameter.STEEL_ELEM_PROFILE_VOLUME: ✅ Fast, ✅ Accurate, ✅ Revit 2022, ✅ Portable
- LookupParameter("Volume"): ✅ Fast, ⚠️ Variable accuracy (detail dependent), ✅ Revit 2022, ❌ Not portable
- Geometry Solid.Volume: ❌ Slow, ⭐ Most accurate (actual geometry), ✅ Revit 2022, ✅ Portable

## Key Issues to Avoid

1. Import-time missing enum members → Use runtime method invocation (Building Coder pattern)
2. Language translation issues → Use BuiltInParameter enum not LookupParameter by name
3. Performance bottlenecks → Avoid get_Geometry() for batch processing
4. Detail level inconsistencies → Use GetMaterialVolume() for consistency
5. Unit conversion errors → Internal units are cubic feet, convert appropriately
