# Add a chartsheet to a workbook

The `wb_add_chartsheet()` function appends a specialized chartsheet to a
`wbWorkbook` object. Unlike standard worksheets, which contain a grid of
cells, a chartsheet is dedicated exclusively to the display of a single,
full-page chart.

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

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object to which the new chartsheet will be attached.

- sheet:

  A character string for the chartsheet name. Defaults to a sequentially
  generated name (e.g., "Sheet 1").

- tab_color:

  The color of the sheet tab. Accepts a
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object, a standard R color name, or a hex color code.

- zoom:

  The zoom level as a percentage; a numeric value between 10 and 400.

- visible:

  The visibility state of the sheet. Options include "visible",
  "hidden", or "veryHidden".

- ...:

  Additional arguments passed to internal configuration methods.

## Details

A chartsheet is a distinct sheet type in the OpenXML specification. It
does not support standard cell data, grid lines, or typical worksheet
features. Its primary purpose is to provide a high-level, focused view
of a graphical representation.

**Important:** A chartsheet must contain a chart object to be valid.
Adding a chartsheet without subsequently attaching a chart via
[`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_mschart.md)
will result in a corrupt workbook that may fail to open in spreadsheet
software.

Like standard worksheets, chartsheets support visual customization such
as `tab_color`, `zoom` levels, and various `visible` states.

## See also

[`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_mschart.md),
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md)

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md),
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
