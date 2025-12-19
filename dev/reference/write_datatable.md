# Defunct: Write to a worksheet as a spreadsheet table

Write to a worksheet and format as a spreadsheet table. Use
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md)
in new code. This function is deprecated and may not be exported in the
future.

## Usage

``` r
write_datatable(
  wb,
  sheet,
  x,
  dims = wb_dims(start_row, start_col),
  start_col = 1,
  start_row = 1,
  col_names = TRUE,
  row_names = FALSE,
  table_style = "TableStyleLight9",
  table_name = NULL,
  with_filter = TRUE,
  sep = ", ",
  first_column = FALSE,
  last_column = FALSE,
  banded_rows = TRUE,
  banded_cols = FALSE,
  apply_cell_style = TRUE,
  remove_cell_style = FALSE,
  na = na_strings(),
  inline_strings = TRUE,
  total_row = FALSE,
  ...
)
```

## Arguments

- wb:

  A Workbook object containing a worksheet.

- sheet:

  The worksheet to write to. Can be the worksheet index or name.

- x:

  A data frame

- dims:

  Spreadsheet cell range that will determine `start_col` and
  `start_row`: "A1", "A1:B2", "A:B"

- start_col:

  A vector specifying the starting column to write `x` to.

- start_row:

  A vector specifying the starting row to write `x` to.

- col_names:

  If `TRUE`, column names of `x` are written.

- row_names:

  If `TRUE`, the row names of `x` are written.

- table_style:

  Any table style name or "none" (see
  [`vignette("openxlsx2_style_manual")`](https://janmarvin.github.io/openxlsx2/dev/articles/openxlsx2_style_manual.md))

- table_name:

  Name of table in workbook. The table name must be unique.

- with_filter:

  If `TRUE`, columns with have filters in the first row.

- sep:

  Only applies to list columns. The separator used to collapse list
  columns to a character vector e.g.
  `sapply(x$list_column, paste, collapse = sep)`.

- first_column:

  logical. If `TRUE`, the first column is bold.

- last_column:

  logical. If `TRUE`, the last column is bold.

- banded_rows:

  logical. If `TRUE`, rows are color banded.

- banded_cols:

  logical. If `TRUE`, the columns are color banded.

- apply_cell_style:

  Should we write cell styles to the workbook

- remove_cell_style:

  keep the cell style?

- na:

  Value used for replacing `NA` values from `x`. Default looks if
  `options("openxlsx2.na")` is set. Otherwise
  [`na_strings()`](https://janmarvin.github.io/openxlsx2/dev/reference/waivers.md)
  uses the special `#N/A` value within the workbook.

- inline_strings:

  write characters as inline strings

- total_row:

  logical. With the default `FALSE` no total row is added.

- ...:

  additional arguments

## Details

Formulae written using
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md)
to a Workbook object will not get picked up by
[`read_xlsx()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_to_df.md).
This is because only the formula is written into the worksheet and it
will be evaluated once the file is opened in spreadsheet software. The
string `"_openxlsx_NA"` is reserved for `openxlsx2`. If the data frame
contains this string, the output will be broken. Similar factor labels
`"_openxlsx_Inf"`, `"_openxlsx_nInf"`, and `"_openxlsx_NaN"` are
reserved. The `na` string `"_openxlsx_NULL"` is a special that will be
treated as NULL. So that setting the option
`options("openxlsx2.na" = "_openxlsx_NULL")` will behave similar to
`na = NULL`.

Supported classes are data frames, matrices and vectors of various types
and everything that can be converted into a data frame with
[`as.data.frame()`](https://rdrr.io/r/base/as.data.frame.html).
Everything else that the user wants to write should either be converted
into a vector or data frame or written in vector or data frame segments.
This includes base classes such as `table`, which were coerced
internally in the predecessor of this package.

Even vectors and data frames can consist of different classes. Many base
classes are covered, though not all and far from all third-party
classes. When data of an unknown class is written, it is handled with
[`as.character()`](https://rdrr.io/r/base/character.html). It is not
possible to write character nodes beginning with `<r>` or `<r/>`. Both
are reserved for internal functions. If you need these. You have to wrap
the input string in
[`fmt_txt()`](https://janmarvin.github.io/openxlsx2/dev/reference/fmt_txt.md).

The columns of `x` with class Date/POSIXt, currency, accounting,
hyperlink, percentage are automatically styled as dates, currency,
accounting, hyperlinks, percentages respectively. When writing POSIXt,
the users local timezone should not matter. The openxml standard does
not have a timezone and the conversion from the local timezone should
happen internally, so that date and time are converted, but the timezone
is dropped. This conversion could cause a minor precision loss. The
datetime in R and in spreadsheets might differ by 1 second, caused by
floating point precision. When read from the worksheet, starting with
`openxlsx2` release `1.15` the datetime is returned in `"UTC"`.

Functions
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md)
and
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md)
behave quite similar. The distinction is that the latter creates a table
in the worksheet that can be used for different kind of formulas and can
be sorted independently, though is less flexible than basic cell
regions.
