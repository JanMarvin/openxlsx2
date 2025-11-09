# Add a chartsheet to a workbook

A chartsheet is a special type of sheet that handles charts output. You
must add a chart to the sheet. Otherwise, this will break the workbook.

## Usage

``` r
wb_add_chartsheet(
  wb,
  sheet = next_sheet(),
  tab_color = NULL,
  zoom = 100,
  visible = c("true", "false", "hidden", "visible", "veryhidden"),
  ...
)
```

## Arguments

- wb:

  A Workbook object to attach the new chartsheet

- sheet:

  A name for the new chartsheet

- tab_color:

  Color of the sheet tab. A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md),
  a valid color (belonging to
  [`grDevices::colors()`](https://rdrr.io/r/grDevices/colors.html)) or a
  valid hex color beginning with "#".

- zoom:

  The sheet zoom level, a numeric between 10 and 400 as a percentage. (A
  zoom value smaller than 10 will default to 10.)

- visible:

  If `FALSE`, sheet is hidden else visible.

- ...:

  Additional arguments

## See also

[`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_mschart.md)

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_worksheet.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_workbook.md)
