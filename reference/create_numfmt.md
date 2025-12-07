# Create number format

This function creates a number format for a cell in a spreadsheet.
Number formats define how numeric values are displayed, including dates,
times, currencies, percentages, and more.

## Usage

``` r
create_numfmt(numFmtId = 164, formatCode = "#,##0.00")
```

## Arguments

- numFmtId:

  An ID representing the number format. The list of valid IDs can be
  found in the **Details** section of
  [`create_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/create_cell_style.md).

- formatCode:

  A format code that specifies the display format for numbers. This can
  include custom formats for dates, times, and other numeric values.

## Value

A formatted number format object to be used in a spreadsheet.

## See also

[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_numfmt.md)

Other style creating functions:
[`create_border()`](https://janmarvin.github.io/openxlsx2/reference/create_border.md),
[`create_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/create_cell_style.md),
[`create_colors_xml()`](https://janmarvin.github.io/openxlsx2/reference/create_colors_xml.md),
[`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md),
[`create_fill()`](https://janmarvin.github.io/openxlsx2/reference/create_fill.md),
[`create_font()`](https://janmarvin.github.io/openxlsx2/reference/create_font.md),
[`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/reference/create_tablestyle.md)

## Examples

``` r
# Create a number format for currency
numfmt <- create_numfmt(
  numFmtId = 164,
  formatCode = "$#,##0.00"
)
```
