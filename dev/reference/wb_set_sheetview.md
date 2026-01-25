# Modify the default view of a worksheet

The `wb_set_sheetview()` function controls the visual presentation of a
worksheet. It allows you to toggle UI elements like grid lines,
row/column headers, and formula visibility, as well as setting the zoom
level and view mode (e.g., Normal vs. Page Layout).

## Usage

``` r
wb_set_sheetview(
  wb,
  sheet = current_sheet(),
  color_id = NULL,
  default_grid_color = NULL,
  right_to_left = NULL,
  show_formulas = NULL,
  show_grid_lines = NULL,
  show_outline_symbols = NULL,
  show_row_col_headers = NULL,
  show_ruler = NULL,
  show_white_space = NULL,
  show_zeros = NULL,
  tab_selected = NULL,
  top_left_cell = NULL,
  view = NULL,
  window_protection = NULL,
  workbook_view_id = NULL,
  zoom_scale = NULL,
  zoom_scale_normal = NULL,
  zoom_scale_page_layout_view = NULL,
  zoom_scale_sheet_layout_view = NULL,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet. Defaults to the current sheet.

- color_id, default_grid_color:

  Integer; internal color index for grid lines. Default is 64
  (automatic).

- right_to_left:

  Logical; if `TRUE`, column ordering is right-to-left.

- show_formulas:

  Logical; if `TRUE`, cells display their formulas instead of calculated
  values.

- show_grid_lines:

  Logical; if `TRUE` (default), the worksheet grid lines are visible.

- show_outline_symbols:

  Logical; if `TRUE`, shows symbols for grouped rows or columns.

- show_row_col_headers:

  Logical; if `TRUE`, shows the letters (columns) and numbers (rows) at
  the edges of the sheet.

- show_ruler:

  Logical; if `TRUE`, a ruler is shown in "Page Layout" view.

- show_white_space:

  Logical; if `TRUE`, margins and page gaps are shown in "Page Layout"
  view.

- show_zeros:

  Logical; if `FALSE`, cells containing a value of zero appear blank.

- tab_selected:

  Integer; a zero-based index indicating if this sheet tab is selected.

- top_left_cell:

  Character; the address of the cell that should be positioned in the
  top-left corner of the view (e.g., "B10").

- view:

  Character; the view mode. One of `"normal"`, `"pageBreakPreview"`, or
  `"pageLayout"`.

- window_protection:

  Logical; if `TRUE`, the panes within the sheet view are protected.

- workbook_view_id:

  Integer; links the sheet view to a specific global workbook view.

- zoom_scale, zoom_scale_normal, zoom_scale_page_layout_view,
  zoom_scale_sheet_layout_view:

  Integer; the zoom percentage (between 10 and 400).

- ...:

  Additional arguments.

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
object, invisibly.

The `wbWorkbook` object, invisibly

## Details

Sheet views are saved properties that tell the spreadsheet application
how to render the sheet upon opening. These settings are specific to the
worksheet and do not affect the actual data or styles of the cells.

Common Use Cases:

- Zooming: Use `zoom_scale` to make large datasets more readable or to
  provide a high-level dashboard view.

- Clean Layouts: For reports or dashboards, setting
  `show_grid_lines = FALSE` and `show_row_col_headers = FALSE` creates a
  cleaner, application-like interface.

- Audit Mode: Setting `show_formulas = TRUE` is helpful for debugging
  complex spreadsheets by displaying the formulas directly in the cells.

- Right-to-Left: Essential for spreadsheets in languages like Arabic or
  Hebrew.

## Examples

``` r
wb <- wb_workbook()$add_worksheet()

wb$set_sheetview(
  zoom_scale = 75,
  right_to_left = FALSE,
  show_formulas = TRUE,
  show_grid_lines = TRUE,
  show_outline_symbols = FALSE,
  show_row_col_headers = TRUE,
  show_ruler = TRUE,
  show_white_space = FALSE,
  tab_selected = 1,
  top_left_cell = "B1",
  view = "normal",
  window_protection = TRUE
)
```
