# Add a pivot table to a worksheet

The data must be specified using
[`wb_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_data.md)
to ensure the function works. The sheet will be empty unless it is
opened in spreadsheet software. Find more details in the [section about
pivot
tables](https://janmarvin.github.io/ox2-book/chapters/openxlsx2_pivot_tables.html)
in the openxlsx2 book.

## Usage

``` r
wb_add_pivot_table(
  wb,
  x,
  sheet = next_sheet(),
  dims = "A3",
  filter,
  rows,
  cols,
  data,
  fun,
  params,
  pivot_table,
  slicer,
  timeline
)
```

## Arguments

- wb:

  A Workbook object containing a \#' worksheet.

- x:

  A `data.frame` that inherits the
  [`wb_data`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_data.md)
  class.

- sheet:

  A worksheet containing a \#'

- dims:

  The worksheet cell where the pivot table is placed

- filter:

  The column name(s) of `x` used for filter.

- rows:

  The column name(s) of `x` used as rows

- cols:

  The column names(s) of `x` used as cols

- data:

  The column name(s) of `x` used as data

- fun:

  A vector of functions to be used with `data`. See **Details** for the
  list of available options.

- params:

  A list of parameters to modify pivot table creation. See **Details**
  for available options.

- pivot_table:

  An optional name for the pivot table

- slicer, timeline:

  Any additional column name(s) of `x` used as slicer/timeline

## Details

The pivot table is not actually written to the worksheet, therefore the
cell region has to remain empty. What is written to the workbook is
something like a recipe how the spreadsheet software has to construct
the pivot table when opening the file.

It is possible to add slicers to the pivot table. For this the pivot
table has to be named and the variable used as slicer, must be part of
the selected pivot table names (`cols`, `rows`, `filter`, or `slicer`).
If these criteria are matched, a slicer can be added using
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md).

Be aware that you should always test on a copy if a `param` argument
works with a pivot table. Not only to check if the desired effect
appears, but first and foremost if the file loads. Wildly mixing params
might brick the output file and cause spreadsheet software to crash.

`fun` can be any of `AVERAGE`, `COUNT`, `COUNTA`, `MAX`, `MIN`,
`PRODUCT`, `STDEV`, `STDEVP`, `SUM`, `VAR`, `VARP`.

`show_data_as` can be any of `normal`, `difference`, `percent`,
`percentDiff`, `runTotal`, `percentOfRow`, `percentOfCol`,
`percentOfTotal`, `index`.

It is possible to calculate data fields if the formula is assigned as a
variable name for the field to calculate. This would look like this:
`data = c("am", "disp/cyl" = "New")`

Possible `params` arguments are listed below. Pivot tables accepts more
parameters, but they were either not tested or misbehaved (probably
because we misunderstood how the parameter should be used).

Boolean arguments:

- apply_alignment_formats

- apply_number_formats

- apply_border_formats

- apply_font_formats

- apply_pattern_formats

- apply_width_height_formats

- no_style

- compact

- outline

- compact_data

- row_grand_totals

- col_grand_totals

Table styles accepting character strings:

- auto_format_id: style id as character in the range of 4096 to 4117

- table_style: a predefined (pivot) table style `"TableStyleMedium23"`

- show_data_as: accepts character strings as listed above

Miscellaneous:

- numfmt: accepts vectors of the form `c(formatCode = "0.0%")`

- choose: select variables in the form of a named logical vector like
  `c(agegp = 'x > "25-34"')` for the `esoph` dataset.

- sort_item: named list of index or character vectors

## See also

[`wb_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_data.md)

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chartsheet.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
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

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`filter-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/filter-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`named_region-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/named_region-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_conditional_formatting.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_thread.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md)

## Examples

``` r
wb <- wb_workbook()
wb <- wb_add_worksheet(wb)
wb <- wb_add_data(wb, x = mtcars)

df <- wb_data(wb, sheet = 1)

# default pivot table
wb <- wb_add_pivot_table(wb, x = df, dims = "A3",
    filter = "am", rows = "cyl", cols = "gear", data = "disp"
  )
  # with parameters
wb <- wb_add_pivot_table(wb, x = df,
    filter = "am", rows = "cyl", cols = "gear", data = "disp",
    params = list(no_style = TRUE, numfmts = c(formatCode = "##0.0"))
  )
```
