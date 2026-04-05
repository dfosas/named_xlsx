# Welcome to `named_xlsx`

`named_xlsx` provides a small set of workbook-oriented tools for working
with Excel through named cells, including named tables.

## Backends
The package supports the following backends to interact with Excel:

| Engine | Read | Write | Save | Notes |
| --- | --- | --- | --- | --- |
| `OpenPYXL` | Yes | Yes | Yes | Default backend. |
| `XLWings` | Yes | Yes | Yes | Requires Excel and the `xlsx` extra. |
| `Calamine` | Yes | No | No | Read-only backend, install via the `calamine` extra. |


## Table convention
Named tables are expected to be part of the workbook's defined names whose
reference text is a structured table reference such as `tbl_demo[value]`.
For example, a defined name `series` can point to `tbl_demo[value]`.

When resolved by `named_xlsx`, the package returns the table data rows
only, excluding the header and total rows.


## Command-line tools

- `named_xlsx save`: export named-cell values to TOML
- `named_xlsx load`: apply TOML values to a copied workbook
- `named_xlsx spec`: list named-cell coordinates and values
- `named_xlsx refresh`: refresh workbook cached values through reloading with `xlwings`.

The CLI functions accept an `engine` argument.
Read-only commands can use`Calamine`; 
write commands require a writable backend such as `OpenPYXL` or `XLWings`.
