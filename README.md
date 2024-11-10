# Gnumeric-py

A python library for reading and writing [Gnumeric](http://www.gnumeric.org/) files.

## Quick Start

Open an existing workbook and explore what's available:

```python
from gnumeric.workbook import Workbook

wb = Workbook.load_workbook('samples/test.gnumeric')

# Get the active sheet
ws = wb.get_active_sheet()  # can also use wb.active

# Get a sheet by name
ws_1 = wb.get_sheet_by_name('Sheet1')

# Get cell at A1
cell_a1 = ws['A1']
cell_a1 = ws[(0, 0)]  # zero-indexed (row, column)

# Get the value in the cell
cell_a1.get_value()
```

## Unsupported Features

Currently, this library does not support:

1. Evaluating expressions
2. Graphs
