# Defunct: Write an object to a worksheet

Use
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md)
or
[`write_xlsx()`](https://janmarvin.github.io/openxlsx2/dev/reference/write_xlsx.md)
in new code.

## Usage

``` r
write_data(
  wb,
  sheet,
  x,
  dims = wb_dims(start_row, start_col),
  start_col = 1,
  start_row = 1,
  array = FALSE,
  col_names = TRUE,
  row_names = FALSE,
  with_filter = FALSE,
  sep = ", ",
  name = NULL,
  apply_cell_style = TRUE,
  remove_cell_style = FALSE,
  na = na_strings(),
  inline_strings = TRUE,
  enforce = FALSE,
  ...
)
```

## Arguments

- wb:

  A Workbook object containing a worksheet.

- sheet:

  The worksheet to write to. Can be the worksheet index or name.

- x:

  Object to be written. For classes supported look at the examples.

- dims:

  Spreadsheet cell range that will determine `start_col` and
  `start_row`: "A1", "A1:B2", "A:B"

- start_col:

  A vector specifying the starting column to write `x` to.

- start_row:

  A vector specifying the starting row to write `x` to.

- array:

  A bool if the function written is of type array

- col_names:

  If `TRUE`, column names of `x` are written.

- row_names:

  If `TRUE`, the row names of `x` are written.

- with_filter:

  If `TRUE`, add filters to the column name row. NOTE: can only have one
  filter per worksheet.

- sep:

  Only applies to list columns. The separator used to collapse list
  columns to a character vector e.g.
  `sapply(x$list_column, paste, collapse = sep)`.

- name:

  The name of a named region if specified.

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

- enforce:

  enforce that selected dims is filled. For this to work, `dims` must
  match `x`

- ...:

  additional arguments

## Value

invisible(0)
