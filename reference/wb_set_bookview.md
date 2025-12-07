# Get and Set the workbook position, size and filter

Get and Set the workbook position, size and filter

## Usage

``` r
wb_get_bookview(wb)

wb_remove_bookview(wb, view = NULL)

wb_set_bookview(
  wb,
  active_tab = NULL,
  auto_filter_date_grouping = NULL,
  first_sheet = NULL,
  minimized = NULL,
  show_horizontal_scroll = NULL,
  show_sheet_tabs = NULL,
  show_vertical_scroll = NULL,
  tab_ratio = NULL,
  visibility = NULL,
  window_height = NULL,
  window_width = NULL,
  x_window = NULL,
  y_window = NULL,
  view = 1L,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object

- view:

  Which view to modify. Default is `1` (the first view).

- active_tab:

  activeTab

- auto_filter_date_grouping:

  autoFilterDateGrouping

- first_sheet:

  The first sheet to be displayed

- minimized:

  minimized

- show_horizontal_scroll:

  showHorizontalScroll

- show_sheet_tabs:

  showSheetTabs

- show_vertical_scroll:

  showVerticalScroll

- tab_ratio:

  tabRatio

- visibility:

  visibility

- window_height:

  windowHeight

- window_width:

  windowWidth

- x_window:

  xWindow

- y_window:

  yWindow

- ...:

  additional arguments

## Value

A data frame with the bookview properties

The Workbook object

The Workbook object

## Examples

``` r
 wb <- wb_workbook()
 wb <- wb_add_worksheet(wb)

 # set the first and second bookview (horizontal split)
 wb <- wb_set_bookview(wb,
     window_height = 17600, window_width = 15120,
     x_window = 15120, y_window = 760)
 wb <- wb_set_bookview(wb,
     window_height = 17600, window_width = 15040,
     x_window = 0, y_window = 760, view = 2
   )

 wb_get_bookview(wb)
#>   windowHeight windowWidth xWindow yWindow
#> 1        17600       15120   15120     760
#> 2        17600       15040       0     760

 # remove the first view
 wb <- wb_remove_bookview(wb, view = 1)
 wb_get_bookview(wb)
#>   windowHeight windowWidth xWindow yWindow
#> 1        17600       15040       0     760

 # keep only the first view
 wb <- wb_remove_bookview(wb, view = -1)
 wb_get_bookview(wb)
#>   windowHeight windowWidth xWindow yWindow
#> 1        17600       15040       0     760
```
