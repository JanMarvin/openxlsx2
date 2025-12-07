# Set the default font in a workbook

Modify / get the default font for the workbook. This will alter the
latin major and minor font in the workbooks theme.

## Usage

``` r
wb_set_base_font(
  wb,
  font_size = 11,
  font_color = wb_color(theme = "1"),
  font_name = "Aptos Narrow",
  ...
)

wb_get_base_font(wb)
```

## Arguments

- wb:

  A workbook object

- font_size:

  Font size

- font_color:

  Font color

- font_name:

  Name of a font

- ...:

  Additional arguments

## Details

The font name is not validated in anyway. Spreadsheet software replaces
unknown font names with system defaults.

The default base font is Aptos Narrow, black, size 11. If `font_name`
differs from the name in `wb_get_base_font()`, the theme is updated to
use the newly selected font name.

## See also

Other workbook styling functions:
[`wb_add_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_dxfs_style.md),
[`wb_add_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_style.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md)

Other workbook wrappers:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md),
[`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_chartsheet.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md),
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

## Examples

``` r
## create a workbook
wb <- wb_workbook(theme = "Office 2013 - 2022 Theme")
wb$add_worksheet("S1")
## modify base font to size 10 Aptos Narrow in red
wb$set_base_font(font_size = 10, font_color = wb_color("red"), font_name = "Aptos Narrow")

wb$add_data(x = iris)

## font color does not affect tables
wb$add_data_table(x = iris, dims = wb_dims(from_col = 10))

## get the base font
wb_get_base_font(wb)
#> $size
#> $size$val
#> [1] "10"
#> 
#> 
#> $color
#> $color$rgb
#> [1] "FFFF0000"
#> 
#> 
#> $name
#> $name$val
#> [1] "Aptos Narrow"
#> 
#> 
```
