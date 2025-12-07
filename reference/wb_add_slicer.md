# Add a slicer/timeline to a pivot table

Add a slicer/timeline to a previously created pivot table. This function
is still experimental and might be changed/improved in upcoming
releases.

## Usage

``` r
wb_add_slicer(
  wb,
  x,
  dims = "A1",
  sheet = current_sheet(),
  pivot_table,
  slicer,
  params
)

wb_remove_slicer(wb, sheet = current_sheet())

wb_add_timeline(
  wb,
  x,
  dims = "A1",
  sheet = current_sheet(),
  pivot_table,
  timeline,
  params
)

wb_remove_timeline(wb, sheet = current_sheet())
```

## Arguments

- wb:

  A Workbook object containing a worksheet.

- x:

  A `data.frame` that inherits the
  [`wb_data`](https://janmarvin.github.io/openxlsx2/reference/wb_data.md)
  class.

- dims:

  The worksheet cell where the pivot table is placed

- sheet:

  A worksheet

- pivot_table:

  The name of a pivot table

- slicer, timeline:

  A variable used as slicer/timeline for the pivot table

- params:

  A list of parameters to modify pivot table creation. See **Details**
  for available options.

## Details

This assumes that the slicer/timeline variable initialization has
happened before. Unfortunately, it is unlikely that we can guarantee
this for loaded workbooks, and we *strictly* discourage users from
attempting this. If the variable has not been initialized properly, this
may cause the spreadsheet software to crash. Although it is documented
that slicers should use "TimelineStyleLight\[1-6\]" and
"TimelineStyleDark\[1-6\]" they use slicer styles.

Possible `params` arguments for slicers are listed below.

- edit_as: "twoCell" to place the slicer into the cells

- column_count: integer used as column count

- sort_order: "descending" / "ascending"

- choose: select variables in the form of a named logical vector like
  `c(agegp = 'x > "25-34"')` for the `esoph` dataset.

- locked_position

- start_item

- hide_no_data_items

Possible `params` arguments for timelines are listed below.

- beg_date/end_date: dates when the timeline should begin or end

- choose_beg/choose_end: dates when the selection should begin or end

- scroll_position

- show_selection_label

- show_time_level

- show_horizontal_scrollbar

Possible common `params`:

- caption: string used for a caption

- style: "SlicerStyleLight\[1-6\]", "SlicerStyleDark\[1-6\]" only for
  slicer "SlicerStyleOther\[1-2\]"

- level: the granularity of the slicer (for timeline 0 = year, 1 =
  quarter, 2 = month)

- show_caption: logical if caption should be shown or not

Removing works on the spreadsheet level. Therefore all slicers/timelines
are removed from a worksheet. At the moment the drawing reference
remains on the spreadsheet. Therefore spreadsheet software that does not
handle slicers/timelines will still show the drawing.

## See also

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md),
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
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
[`filter-wb`](https://janmarvin.github.io/openxlsx2/reference/filter-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
[`named_region-wb`](https://janmarvin.github.io/openxlsx2/reference/named_region-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md),
[`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_conditional_formatting.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md)

## Examples

``` r
# prepare data
df <- data.frame(
  AirPassengers = c(AirPassengers),
  time = seq(from = as.Date("1949-01-01"), to = as.Date("1960-12-01"), by = "month"),
  letters = letters[1:4],
  stringsAsFactors = FALSE
)

# create workbook
wb <- wb_workbook()$
  add_worksheet("pivot")$
  add_worksheet("data")$
  add_data(x = df)

# get pivot table data source
df <- wb_data(wb, sheet = "data")

# create pivot table
wb$add_pivot_table(
  df,
  sheet = "pivot",
  rows = "time",
  cols = "letters",
  data = "AirPassengers",
  pivot_table = "airpassengers",
  params = list(
    compact = FALSE, outline = FALSE, compact_data = FALSE,
    row_grand_totals = FALSE, col_grand_totals = FALSE)
)

# add slicer
wb$add_slicer(
  df,
  dims = "E1:I7",
  sheet = "pivot",
  slicer = "letters",
  pivot_table = "airpassengers",
  params = list(choose = c(letters = 'x %in% c("a", "b")'))
)

# add timeline
wb$add_timeline(
  df,
  dims = "E9:I14",
  sheet = "pivot",
  timeline = "time",
  pivot_table = "airpassengers",
  params = list(
    beg_date = as.Date("1954-01-01"),
    end_date = as.Date("1961-01-01"),
    choose_beg = as.Date("1957-01-01"),
    choose_end = as.Date("1958-01-01"),
    level = 0,
    style = "TimeSlicerStyleLight2"
  )
)
```
