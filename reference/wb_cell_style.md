# Get or set cell style indices

The `wb_get_cell_style()` and `wb_set_cell_style()` functions provide a
direct way to manage the internal style index (XF ID) of a cell. This is
particularly useful for "copy-pasting" the formatting from one cell to
another or for applying pre-defined styles at scale without the overhead
of creating new XML nodes for every cell.

## Usage

``` r
wb_get_cell_style(wb, sheet = current_sheet(), dims)

wb_set_cell_style(wb, sheet = current_sheet(), dims, style)

wb_set_cell_style_across(
  wb,
  sheet = current_sheet(),
  style,
  cols = NULL,
  rows = NULL
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet. Defaults to the current sheet.

- dims:

  A character string defining the cell range (e.g., "A1" or "A1:C10").

- style:

  A style or a cell with a certain style

- cols:

  The columns the style will be applied to, either "A:D" or 1:4

- rows:

  The rows the style will be applied to

## Value

- For `wb_get_cell_style()`: A named vector where names are cell
  addresses (e.g., "A1") and values are the integer style indices.

- For `wb_set_cell_style()`: The
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object, invisibly.

A named vector with cell style index positions

## Details

In the OpenXML format, formatting is not stored inside every cell.
Instead, a workbook maintains a centralized style table, and each cell
simply holds an integer index (the Cell Style ID) pointing to a record
in that table.

`wb_get_cell_style()` retrieves these indices for a specified range. If
a cell has not been explicitly styled, the function returns the index
for the default style.

`wb_set_cell_style()` applies a specific index or style definition to a
range. This is significantly faster and more memory-efficient than using
high-level wrappers like
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_font.md)
when applying the exact same style to thousands of individual cells.

## Notes

- These functions are the most efficient way to handle repetitive
  styling tasks in large worksheets.

- If `style` is a character string that is not a cell dimension, it is
  looked up in the workbook's Style Manager.

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_border.md),
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_cell_style.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_fill.md),
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_font.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_named_style.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_numfmt.md)

## Examples

``` r
# set a style in B1
wb <- wb_workbook()$add_worksheet()$
  add_numfmt(dims = "B1", numfmt = "#,0")

# get style from B1 to assign it to A1
numfmt <- wb$get_cell_style(dims = "B1")

# assign style to a1
wb$set_cell_style(dims = "A1", style = numfmt)

# set style across a workbook
wb <- wb_workbook()
wb <- wb_add_worksheet(wb)
wb <- wb_add_fill(wb, dims = "C3", color = wb_color("yellow"))
wb <- wb_set_cell_style_across(wb, style = "C3", cols = "C:D", rows = 3:4)
```
