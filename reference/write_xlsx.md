# Write data to an xlsx file

Write a data frame or list of data frames to an xlsx file.

## Usage

``` r
write_xlsx(x, file, as_table = FALSE, ...)
```

## Arguments

- x:

  An object or a list of objects that can be handled by
  [`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md)
  to write to file.

- file:

  An optional xlsx file name. If no file is passed, the object is not
  written to disk and only a workbook object is returned.

- as_table:

  If `TRUE`, will write as a data table, instead of data.

- ...:

  Arguments passed on to
  [`wb_workbook`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md),
  [`wb_add_worksheet`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md),
  [`wb_add_data_table`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md),
  [`wb_add_data`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
  [`wb_freeze_pane`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
  [`wb_set_col_widths`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
  [`wb_save`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
  [`wb_set_base_font`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md)

  `creator`

  :   Creator of the workbook (your name). Defaults to login username or
      `options("openxlsx2.creator")` if set.

  `sheet`

  :   A character string for the worksheet name. Defaults to a
      sequentially generated name (e.g., "Sheet 1").

  `grid_lines`

  :   Logical; if `FALSE`, the worksheet grid lines are hidden.

  `tab_color`

  :   The color of the worksheet tab. Accepts a
      [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
      object, a standard R color name, or a hex color code (e.g.,
      "#4F81BD").

  `zoom`

  :   The sheet zoom level as a percentage; a numeric value between 10
      and 400. Values below 10 default to 10.

  `total_row`

  :   logical. With the default `FALSE` no total row is added.

  `start_col`

  :   A vector specifying the starting column to write `x` to.

  `start_row`

  :   A vector specifying the starting row to write `x` to.

  `col_names`

  :   If `TRUE`, column names of `x` are written.

  `row_names`

  :   If `TRUE`, the row names of `x` are written.

  `na`

  :   Value used for replacing `NA` values from `x`. Default looks if
      `options("openxlsx2.na")` is set. Otherwise
      [`na_strings()`](https://janmarvin.github.io/openxlsx2/reference/waivers.md)
      uses the special `#N/A` value within the workbook.

  `first_active_row`

  :   The index of the first row that should remain scrollable. Rows
      above this will be frozen.

  `first_active_col`

  :   The index or character label of the first column that should
      remain scrollable. Columns to the left will be frozen.

  `first_row`

  :   Logical; if `TRUE`, freezes the first row of the worksheet.

  `first_col`

  :   Logical; if `TRUE`, freezes the first column of the worksheet.

  `widths`

  :   Width to set `cols` to specified column width or `"auto"` for
      automatic sizing. `widths` is recycled to the length of `cols`.
      openxlsx2 sets the default width is 8.43, as this is the standard
      in some spreadsheet software. See **Details** for general
      information on column widths.

  `overwrite`

  :   If `FALSE`, will not overwrite when `file` already exists.

  `font_size`

  :   Font size

  `font_color`

  :   Font color

  `font_name`

  :   Name of a font

## Value

A workbook object

## Details

columns of `x` with class `Date` or `POSIXt` are automatically styled as
dates and datetimes respectively.

## Examples

``` r
## write to working directory
write_xlsx(iris, file = temp_xlsx(), col_names = TRUE)

write_xlsx(iris,
  file = temp_xlsx(),
  col_names = TRUE
)

## Lists elements are written to individual worksheets, using list names as sheet names if available
l <- list("IRIS" = iris, "MTCARS" = mtcars, matrix(runif(1000), ncol = 5))
write_xlsx(l, temp_xlsx(), col_widths = c(NA, "auto", "auto"))

## different sheets can be given different parameters
write_xlsx(l, temp_xlsx(),
  start_col = c(1, 2, 3), start_row = 2,
  as_table = c(TRUE, TRUE, FALSE), with_filter = c(TRUE, FALSE, FALSE)
)

# specify column widths for multiple sheets
write_xlsx(l, temp_xlsx(), col_widths = 20)
write_xlsx(l, temp_xlsx(), col_widths = list(100, 200, 300))
write_xlsx(l, temp_xlsx(), col_widths = list(rep(10, 5), rep(8, 11), rep(5, 5)))

# set base font color to automatic so LibreOffice dark mode works as expected
write_xlsx(l, temp_xlsx(), font_color = wb_color(auto = TRUE))
```
