# SKILL.md — pyRevit/IronPython Extension Development

## Domain

This skill covers development of pyRevit extensions for Autodesk Revit using IronPython. Scripts execute inside Revit's hosted IronPython runtime — they cannot be run, tested, or linted locally.

## IronPython Constraints

IronPython is based on Python 2.7 with partial Python 3 backports via `__future__`. Critical restrictions:

- **No f-strings** — use `.format()` exclusively
- **No `pathlib`** — use `os.path`
- **No `typing` module** — no type hints
- **No `asyncio`**, no `dataclasses`, no `walrus operator`
- `unicode()` exists (Python 2 str/unicode split). Always use `u""` prefix for non-ASCII
- `basestring` may not exist — provide fallback:
  ```python
  try:
      STRING_TYPES = (basestring,)
  except NameError:
      STRING_TYPES = (str,)
  ```

## Revit API Essentials

### Transactions

Every modification requires a `Transaction`:
```python
t = DB.Transaction(doc, u"Tool Name")
t.Start()
try:
    # modifications here
    t.Commit()
except Exception:
    try:
        if t.GetStatus() == DB.TransactionStatus.Started:
            t.RollBack()
    except Exception:
        pass
```

### Element Collection

```python
elements = (
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_Walls)
    .WhereElementIsNotElementType()
    .ToElements()
)
```

### Parameter Access

```python
# Instance parameter
p = elem.LookupParameter(u"ADSK_Размер_Высота")

# Type parameter
elem_type = doc.GetElement(elem.GetTypeId())
p = elem_type.LookupParameter(u"Толщина")

# Always guard with try/except and check None + HasValue
```

### Storage Types

| StorageType | Read | Write |
|---|---|---|
| `String` | `.AsString()` | `.Set(u"value")` |
| `Double` | `.AsDouble()` | `.Set(float_val)` |
| `Integer` | `.AsInteger()` | `.Set(int_val)` |
| `ElementId` | `.AsElementId()` | `.Set(elem_id)` |

Display value (with units): `.AsValueString()` — read-only, for display/reporting.

## pyRevit Framework

### Standard Script Boilerplate

```python
# -*- coding: utf-8 -*-
from __future__ import print_function, division
from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()

def main():
    pass

if __name__ == u"__main__":
    main()
```

### Output Window

```python
output = script.get_output()

output.print_md(u"## Heading")
output.print_md(u"> Summary")
output.print_table(
    table_data=[[u"a", u"b"], [u"c", u"d"]],
    columns=[u"Col1", u"Col2"],
    formats=[u"{}", u"{}"],
)
output.linkify(DB.ElementId(123))  # clickable element link in Revit
```

### User Interaction

```python
forms.alert(u"Message", title=u"Title")
selected = forms.SelectFromList.show(
    items, title=u"Select", multiselect=True
)
```

## Project Conventions

### File Encoding

Every file starts with `# -*- coding: utf-8 -*-`. All Russian text uses `u""` prefix.

### Import Order

1. `__future__`
2. Standard library (`os`, `sys`, `re`, etc.)
3. pyRevit (`from pyrevit import ...`)
4. Revit API (`from Autodesk.Revit.DB import *`)
5. .NET/System (`from System import Enum`)
6. Local modules (after `sys.path.insert(0, lib_dir)`)

### Naming

| Element | Convention | Example |
|---|---|---|
| Files | `lower_snake_case.py` | `nwc_export_utils.py` |
| Entry scripts | `*_script.py` | `DGP_script.py` |
| Functions | `snake_case` | `get_param()`, `to_unicode()` |
| Constants | `UPPER_SNAKE_CASE` | `PARAM_TARGET_HEIGHT` |
| Dict keys | `u"snake_case"` | `{u"status": u"updated"}` |
| Categories (Russian) | via `cat_obj.Name` | `DB.Category.GetCategory(doc, bic).Name` |

### Error Handling Philosophy

Scripts process batches of elements. A single failure must not crash the script:

```python
for elem in elements:
    try:
        # process elem
    except Exception:
        # log, continue to next elem
```

All Revit API property accesses (`elem.Id`, `elem.Category`, etc.) must be wrapped in try/except — they can throw on deleted, invalid, or non-matching elements.

### Report Pattern

Processing functions collect statistics in a dict, then a separate `print_report(stats)` function renders output. Processing logic and presentation are strictly separated.

### Rule-Based Parameter Assignment

Parameters are filled via a RULES list of dicts:
```python
{
    u"category": DB.BuiltInCategory.OST_Walls,
    u"group": u"Фасад",                    # optional condition
    u"groups": [u"Дверь", u"Ворота"],       # or multiple conditions
    u"assignments": [
        {u"target": u"ДГП_Высота", u"source": u"ADSK_Размер_Высота"},
        {u"target": u"ДГП_Толщина", u"source": u"type:Толщина"},
        {u"target": u"ДГП_Код", u"source": u"const:9999"},
        {u"target": u"ДГП_Зона", u"source": u"cond:Назначение=Квартира?Жилая зона:Нежилая зона"},
    ],
}
```

Source spec formats:
- `PARAM_NAME` — read from instance parameter
- `type:PARAM_NAME` — read from type parameter
- `const:VALUE` — literal constant
- `cond:Param=Value?Yes:No` — conditional (empty source = error)

### Bilingual Code

- Code identifiers: **English** (`get_param`, `elem_id`, `process_elements`)
- UI text, parameter names, comments, reports: **Russian** (`u"ДГП_Высота"`, `u"Обработано"`)
- Do not add comments unless explicitly requested
