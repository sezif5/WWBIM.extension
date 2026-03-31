# AGENTS.md

Guide for agentic coding agents working in this repository.

## Project Overview

pyRevit extension for Autodesk Revit. IronPython runtime. 78+ Python scripts organized as pyRevit UI buttons (`*.pushbutton/`), shared libraries (`lib/`), and batch operations.

**There is no build step, no package manager, no test framework, no CI/CD, and no linter configured.** Scripts run inside Revit via pyRevit. You cannot execute or test them locally.

## Build/Lint/Test Commands

```bash
# None available. No build, lint, or test commands exist.
# No pytest, ruff, flake8, mypy, or any tooling is configured.
# Scripts execute only inside Revit via pyRevit runtime.
# Verify syntax only: python -m py_compile <file.py>
```

## Architecture

- `WWBIM.extension/` — pyRevit extension root
- `WWBIM.extension/lib/` — shared helpers (`openbg.py`, `closebg.py`, `nwc_export_utils.py`, etc.)
- `WWBIM.extension/lib/Batch Operations/` — batch parameter fill scripts
- `WWBIM.extension/WW.BIM.tab/` — pyRevit UI layout with panels, stacks, pulldowns, and pushbuttons
- Each `.pushbutton/` folder contains: `bundle.yaml`, `<name>_script.py` (entry point), `icon.png`, optional `README.md`

## Code Style

### Encoding

Every file must start with:
```python
# -*- coding: utf-8 -*-
```

### Imports (order matters)

```python
from __future__ import print_function, division  # if needed

import os
import sys

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import *
from System import Enum
```

1. `__future__` imports
2. Standard library
3. pyRevit framework (`pyrevit.*`)
4. Revit API (`Autodesk.Revit.*`)
5. .NET/System (`System.*`)
6. Local modules (after `sys.path` setup if needed)

### String Formatting

**Always use `.format()`**, never f-strings. This is IronPython:
```python
# Correct
output.print_md(u"Всего: **{}**".format(count))

# Wrong
output.print_md(f"Всего: **{count}**")
```

### Unicode Strings

Use `u""` prefix for all non-ASCII strings (Russian text, parameter names):
```python
PARAM_TARGET_HEIGHT = u"ДГП_Высота"
cat_name = to_unicode(elem.Category.Name)
```

### Naming Conventions

| Element | Style | Example |
|---------|-------|---------|
| Files | `lower_snake_case.py`, entry scripts end `_script.py` | `DGP_script.py`, `nwc_export_utils.py` |
| Classes | PascalCase | `DialogSuppressor`, `ElementCache` |
| Functions | snake_case | `to_unicode()`, `get_param()` |
| Constants | UPPER_SNAKE_CASE | `PARAM_TARGET_HEIGHT`, `DLL_PATHS` |
| Dict keys | `u"snake_case"` with u-prefix | `{u"status": u"updated"}` |
| Variables | snake_case | `cat_name`, `elem_id`, `target_param` |

### Error Handling

Wrap **every** Revit API call in `try/except`. Scripts must never crash — they process batches and continue on failure:
```python
try:
    elem_id = elem.Id.IntegerValue
except Exception:
    elem_id = -1
```

Transaction pattern:
```python
t = DB.Transaction(doc, u"Description")
t.Start()
try:
    # do work
    t.Commit()
except Exception as e:
    try:
        if t.GetStatus() == DB.TransactionStatus.Started:
            t.RollBack()
    except Exception:
        pass
```

### Comments and Docstrings

- Comments in **Russian**
- Section separators: `# ============ SECTION NAME ===========`
- Module docstrings: Russian, triple-quoted, describe purpose
- Function docstrings: Google-style with `Args:` / `Returns:` sections
- **Do not add comments** unless explicitly asked

### Language

- Code identifiers, variable names, function names: **English**
- UI text, parameter names, comments, docstrings, log messages: **Russian**
- Report output: Russian headings with `output.print_md()`, `output.print_table()`

### Dict Return Values

Functions that process elements return result dicts:
```python
{
    u"status": u"updated",       # or "skipped", "already_ok", "exception"
    u"reason": None,             # or error reason string
    u"target": target_name,
    u"source": source_spec,
}
```

### Special Source Spec Formats

Used in rule-based parameter assignment:
- `const:XXX` — literal constant value
- `type:Параметр` — read from element type, not instance
- `cond:Параметр=Значение?ЕслиДа:ЕслиНет` — conditional logic

## Key Patterns

### pyRevit Script Structure

```python
# -*- coding: utf-8 -*-
from __future__ import print_function, division
from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()

def main():
    t = DB.Transaction(doc, u"Tool Name")
    t.Start()
    try:
        # process elements
        t.Commit()
        # print report
    except Exception as e:
        # rollback, print error
        pass

if __name__ == u"__main__":
    main()
```

### pyRevit Output

```python
output.print_md(u"## Heading")
output.print_md(u"> Summary line")
output.print_table(
    table_data=rows,
    columns=[u"Col1", u"Col2"],
    formats=[u"{}", u"{}"],
)
output.linkify(element_id)  # clickable Revit element link
```

### Parameter Access

```python
def get_param(elem, name):
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None

# Always check None and HasValue
p = get_param(elem, param_name)
if p is None:
    # parameter not found
if not p.HasValue:
    # parameter exists but empty
```

## Do's and Don'ts

- **Do** wrap all Revit API calls in `try/except`
- **Do** use `.format()` for string formatting
- **Do** use `u""` prefix for Russian strings
- **Do** use `normalize_for_compare()` when comparing parameter values
- **Do** separate reporting logic from processing logic
- **Do** keep function signatures simple and return result dicts
- **Don't** use f-strings (IronPython)
- **Don't** add comments unless asked
- **Don't** assume Revit API calls won't throw
- **Don't** introduce new config files or dependencies
- **Don't** use `basestring` without a fallback for Python 3
- **Don't** commit `__pycache__/`, `*.pyc`, or `*.zip` files
