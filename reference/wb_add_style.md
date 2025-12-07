# Set the default style in a workbook

wb wrapper to add style to workbook

## Usage

``` r
wb_add_style(wb, style = NULL, style_name = NULL)
```

## Arguments

- wb:

  A workbook

- style:

  style xml character, created by a `create_*()` function.

- style_name:

  style name used optional argument

## Value

The `wbWorkbook` object, invisibly.

## See also

- [`create_border()`](https://janmarvin.github.io/openxlsx2/reference/create_border.md)

- [`create_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/create_cell_style.md)

- [`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md)

- [`create_fill()`](https://janmarvin.github.io/openxlsx2/reference/create_fill.md)

- [`create_font()`](https://janmarvin.github.io/openxlsx2/reference/create_font.md)

- [`create_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/create_numfmt.md)

Other workbook styling functions:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md),
[`wb_add_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_dxfs_style.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md)

## Examples

``` r
yellow_f <- wb_color(hex = "FF9C6500")
yellow_b <- wb_color(hex = "FFFFEB9C")

yellow <- create_dxfs_style(font_color = yellow_f, bg_fill = yellow_b)
wb <- wb_workbook()
wb <- wb_add_style(wb, yellow)
```
