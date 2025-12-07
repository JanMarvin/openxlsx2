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

  :   A name for the new worksheet

  `grid_lines`

  :   A logical. If `FALSE`, the worksheet grid lines will be hidden.

  `tab_color`

  :   Color of the sheet tab. A
      [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md),
      a valid color (belonging to
      [`grDevices::colors()`](https://rdrr.io/r/grDevices/colors.html))
      or a valid hex color beginning with "#".

  `zoom`

  :   The sheet zoom level, a numeric between 10 and 400 as a
      percentage. (A zoom value smaller than 10 will default to 10.)

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

  `na.strings`

  :   Value used for replacing `NA` values from `x`. Default looks if
      `options(openxlsx2.na.strings)` is set. Otherwise
      [`na_strings()`](https://janmarvin.github.io/openxlsx2/reference/waivers.md)
      uses the special `#N/A` value within the workbook.

  `first_active_row`

  :   Top row of active region

  `first_active_col`

  :   Furthest left column of active region

  `first_row`

  :   If `TRUE`, freezes the first row (equivalent to
      `first_active_row = 2`)

  `first_col`

  :   If `TRUE`, freezes the first column (equivalent to
      `first_active_col = 2`)

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
