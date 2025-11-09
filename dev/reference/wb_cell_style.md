# Apply styling to a cell region

Setting a style across only impacts cells that are not yet part of a
workbook. The effect is similar to setting the cell style for all cells
in a row independently, though much quicker and less memory consuming.

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

  A `wbWorkbook` object

- sheet:

  sheet

- dims:

  A cell range in the worksheet

- style:

  A style or a cell with a certain style

- cols:

  The columns the style will be applied to, either "A:D" or 1:4

- rows:

  The rows the style will be applied to

## Value

A named vector with cell style index positions

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_border.md),
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_cell_style.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_fill.md),
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_font.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_named_style.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_numfmt.md)

## Examples

``` r
# set a style in b1
wb <- wb_workbook()$add_worksheet()$
  add_numfmt(dims = "B1", numfmt = "#,0")

# get style from b1 to assign it to a1
numfmt <- wb$get_cell_style(dims = "B1")

# assign style to a1
wb$set_cell_style(dims = "A1", style = numfmt)

# set style across a workbook
wb <- wb_workbook()
wb <- wb_add_worksheet(wb)
wb <- wb_add_fill(wb, dims = "C3", color = wb_color("yellow"))
wb <- wb_set_cell_style_across(wb, style = "C3", cols = "C:D", rows = 3:4)
```
