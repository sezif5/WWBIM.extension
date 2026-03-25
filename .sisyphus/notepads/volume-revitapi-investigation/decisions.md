# Decisions - Volume API Investigation

## Primary Method Selection

**Decision**: Use `GetMaterialVolume()` as primary approach for batch runner

**Rationale**:
1. ✅ Performance-optimized (no expensive geometry extraction)
2. ✅ Accurate (based on Fine detail level, most precise)
3. ✅ Revit 2022+ compatible (method exists since 2014)
4. ✅ Language-portable (no string-based name lookup)
5. ✅ Works for structural columns (3D families with material quantities)

**Reference**: Autodesk/revit-ifc CategoryUtil.cs#L608 (2026)

## Secondary Method Selection

**Decision**: Use `get_Parameter(BuiltInParameter.STEEL_ELEM_PROFILE_VOLUME)` as fallback for steel framing

**Rationale**:
1. ✅ Fast (parameter lookup, no geometry)
2. ✅ Language-independent (enum approach)
3. ✅ Specifically designed for steel profiles
4. ✅ Revit 2022 compatible
5. ✅ Runtime enum resolution avoids import-time errors

**Reference**: CodeCavePro decompiled BuiltInParameter.cs#L53

## Method to Avoid

**Decision**: DO NOT use `get_Geometry()` + `Solid.Volume` in batch runner

**Rationale**:
1. ❌ Performance impact (geometry extraction is expensive)
2. ❌ Only necessary if exact geometry volumes required
3. ⚠️ Better alternatives available for material quantity purposes
4. ℹ️ Reserve for edge cases where geometry volume differs from material volume

## Method to Use with Caution

**Decision**: Use `LookupParameter("Volume")` only as last resort with null-coalescing

**Rationale**:
1. ⚠️ Returns first match if multiple params exist
2. ❌ NOT language-portable (built-in names translated)
3. ✅ Works as fallback when enum not available
4. ⚠️ Always use `.AsDouble()` with null-coalescing operator

**Implementation**:
```csharp
double volume = element.LookupParameter("Volume")?.AsDouble() ?? 0.0;
```

## Detail Level Awareness

**Decision**: Always assume Fine detail level for volume calculations

**Rationale**:
1. Material:Volume uses Fine detail level (most accurate)
2. Volume parameter uses Coarse/Medium detail (less accurate)
3. Medium detail objects never counted in either calculation
4. GetMaterialVolume() provides consistency with Fine detail

**Reference**: Autodesk Support - Material:Volume vs Volume (2023)

## Unit Conversion Strategy

**Decision**: Handle internal units (cubic feet) explicitly

**Rationale**:
1. GetMaterialVolume() returns cubic feet (internal Revit units)
2. Project units may be metric or imperial
3. Need conversion for user-facing display
4. Use UnitUtils.Convert() for display values

## Runtime Enum Resolution Pattern

**Decision**: Use Building Coder reflection pattern for enum safety

**Rationale**:
1. Avoids import-time missing enum member errors
2. Allows graceful degradation if method not available
3. Works across Revit versions (2022+)
4. Provides null-safe behavior

**Reference**: Jeremy Tammik Building Coder Util.cs#L3569

## Implementation Priority Order

1. **Try**: `GetMaterialVolume(materialId)` - performance-optimized
2. **Fallback**: `get_Parameter(BuiltInParameter.STEEL_ELEM_PROFILE_VOLUME)` - steel-specific
3. **Last Resort**: `LookupParameter("Volume")?.AsDouble()` - generic
4. **Avoid**: `get_Geometry()` - unless geometry volume explicitly required

## Batch Processing Strategy

**Decision**: Process elements in batches, not individually

**Rationale**:
1. GetMaterialIds() is cheap, called once per element
2. Then iterate material IDs to get volumes
3. Minimizes API calls and transaction overhead
4. Improves overall batch runner performance

## Error Handling Strategy

**Decision**: Use null-coalescing operator throughout

**Rationale**:
1. Parameters may not exist on all elements
2. Materials may not be assigned to all families
3. Prevents null reference exceptions
4. Provides zero value as default (safe for sums)
