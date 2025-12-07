# Get / Set worksheet names for a workbook

Gets / Sets the worksheet names for a
[wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
object.

## Usage

``` r
wb_set_sheet_names(wb, old = NULL, new)

wb_get_sheet_names(wb, escape = FALSE)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object

- old:

  The name (or index) of the old sheet name. If `NULL` will assume all
  worksheets are to be renamed.

- new:

  The name of the new sheet

- escape:

  Should the xml special characters be escaped?

## Value

- `set_`: The `wbWorkbook` object.

- `get_`: A named character vector of sheet names in order. The names
  represent the original value of the worksheet prior to any character
  substitutions.

## Details

This only changes the sheet name as shown in spreadsheet software and
will not alter it elsewhere. Not in formulas, chart references, named
regions, pivot tables or anywhere else.
