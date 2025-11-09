# Apply styling to a cell region with a named style

Set the styling to a named style for a cell region. Use
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_cell_style.md)
to style a cell region with custom parameters. A named style is the one
in spreadsheet software, like "Normal", "Warning".

## Usage

``` r
wb_add_named_style(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  name = "Normal",
  font_name = NULL,
  font_size = NULL
)
```

## Arguments

- wb:

  A `wbWorkbook` object

- sheet:

  A worksheet

- dims:

  A cell range

- name:

  The named style name. Builtin styles are `Normal`, `Bad`, `Good`,
  `Neutral`, `Calculation`, `Check Cell`, `Explanatory Text`, `Input`,
  `Linked Cell`, `Note`, `Output`, `Warning Text`, `Heading 1`,
  `Heading 2`, `Heading 3`, `Heading 4`, `Title`, `Total`,
  `$x% - Accent$y` (for x in 20, 40, 60 and y in 1:6), `Accent$y` (for y
  in 1:6), `Comma`, `Comma [0]`, `Currency`, `Currency [0]`, `Per cent`

- font_name, font_size:

  optional else the default of the theme

## Value

The `wbWorkbook`, invisibly

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_border.md),
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_cell_style.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_fill.md),
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_font.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_numfmt.md),
[`wb_cell_style`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_cell_style.md)

## Examples

``` r
wb <- wb_workbook()$add_worksheet()
name <- "Normal"
dims <- "A1"
wb$add_data(dims = dims, x = name)

name <- "Bad"
dims <- "B1"
wb$add_named_style(dims = dims, name = name)
wb$add_data(dims = dims, x = name)

name <- "Good"
dims <- "C1"
wb$add_named_style(dims = dims, name = name)
wb$add_data(dims = dims, x = name)
```
