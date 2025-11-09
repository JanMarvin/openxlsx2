# Defunct: Write a character vector as a spreadsheet Formula

Write a a character vector containing spreadsheet formula to a
worksheet. Use
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md)
or `add_formula()` in new code. This function is deprecated and may be
defunct.

## Usage

``` r
write_formula(
  wb,
  sheet,
  x,
  dims = wb_dims(start_row, start_col),
  start_col = 1,
  start_row = 1,
  array = FALSE,
  cm = FALSE,
  apply_cell_style = TRUE,
  remove_cell_style = FALSE,
  enforce = FALSE,
  ...
)
```

## Arguments

- wb:

  A Workbook object containing a worksheet.

- sheet:

  The worksheet to write to. (either as index or name)

- x:

  A formula as character vector.

- dims:

  Spreadsheet dimensions that will determine where `x` spans: "A1",
  "A1:B2", "A:B"

- start_col:

  A vector specifying the starting column to write to.

- start_row:

  A vector specifying the starting row to write to.

- array:

  A bool if the function written is of type array

- cm:

  A special kind of array function that hides the curly braces in the
  cell. Add this, if you see "@" inserted into your formulas.

- apply_cell_style:

  Should we write cell styles to the workbook?

- remove_cell_style:

  Should we keep the cell style?

- enforce:

  enforce dims

- ...:

  additional arguments
