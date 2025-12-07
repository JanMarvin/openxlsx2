# Modify the default view of a worksheet

This helps set a worksheet's appearance, such as the zoom, whether to
show grid lines

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

  A Workbook object

- sheet:

  sheet

- color_id, default_grid_color:

  Integer: A color, default is 64

- right_to_left:

  Logical: if `TRUE` column ordering is right to left

- show_formulas:

  Logical: if `TRUE` cell formulas are shown

- show_grid_lines:

  Logical: if `TRUE` the worksheet grid is shown

- show_outline_symbols:

  Logical: if `TRUE` outline symbols are shown

- show_row_col_headers:

  Logical: if `TRUE` row and column headers are shown

- show_ruler:

  Logical: if `TRUE` a ruler is shown in page layout view

- show_white_space:

  Logical: if `TRUE` margins are shown in page layout view

- show_zeros:

  Logical: if `FALSE` cells containing zero are shown blank if
  `show_formulas = FALSE`

- tab_selected:

  Integer: zero vector indicating the selected tab

- top_left_cell:

  Cell: the cell shown in the top left corner / or top right with
  rightToLeft

- view:

  View: "normal", "pageBreakPreview" or "pageLayout"

- window_protection:

  Logical: if `TRUE` the panes are protected

- workbook_view_id:

  integer: Pointing to some other view inside the workbook

- zoom_scale, zoom_scale_normal, zoom_scale_page_layout_view,
  zoom_scale_sheet_layout_view:

  Integer: the zoom scale should be between 10 and 400. These are values
  for current, normal etc.

- ...:

  additional arguments

## Value

The `wbWorkbook` object, invisibly

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
