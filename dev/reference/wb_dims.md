# Helper to specify the `dims` argument

`wb_dims()` can be used to help provide the `dims` argument, in the
`wb_add_*` functions. It returns a A1 spreadsheet range ("A1:B1" or
"A2"). It can be very useful as you can specify many parameters that
interact together In general, you must provide named arguments.
`wb_dims()` will only accept unnamed arguments if they are `rows`,
`cols`, for example `wb_dims(1:4, 1:2)`, that will return "A1:B4".

`wb_dims()` can also be used with an object (a `data.frame` or a
`matrix` for example.) All parameters are numeric unless stated
otherwise.

## Usage

``` r
wb_dims(..., select = NULL)
```

## Arguments

- ...:

  construct `dims` arguments, from rows/cols vectors or objects that can
  be coerced to data frame. `x`, `rows`, `cols`, `from_row`, `from_col`,
  `from_dims` `row_names`, and `col_names` are accepted.

- select:

  A string, one of the followings. it improves the selection of various
  parts of `x` One of "x", "data", "col_names", or "row_names". `"data"`
  will only select the data part, excluding row names and column names
  (default if `cols` or `rows` are specified) `"x"` Includes the data,
  column and row names if they are present. (default if none of `rows`
  and `cols` are provided) `"col_names"` will only return column names
  `"row_names"` Will only return row names.

## Value

A `dims` string

## Details

When using `wb_dims()` with an object, the default behavior is to select
only the data / row or columns in `x` If you need another behavior, use
`wb_dims()` without supplying `x`.

- `x` An object (typically a `matrix` or a `data.frame`, but a vector is
  also accepted.)

- `from_row` / `from_col` / `from_dims` the starting position of `x`
  (The `dims` returned will assume that the top left corner of `x` is at
  `from_row / from_col`

- `rows` Optional Which row span in `x` should this apply to. If `rows`
  = 0, only column names will be affected.

- `cols` a range of columns id in `x`, or one of the column names of `x`
  (length 1 only accepted for column names of `x`.)

- `row_names` A logical, this is to let `wb_dims()` know that `x` has
  row names or not. If `row_names = TRUE`, `wb_dims()` will increment
  `from_col` by 1.

- `col_names` `wb_dims()` assumes that if `x` has column names, then
  trying to find the `dims`.

`wb_dims()` tries to support most possible cases with `row_names = TRUE`
and `col_names = FALSE`, but it works best if `x` has named dimensions
(`data.frame`, `matrix`), and those parameters are not specified. data
with column names, and without row names. as the code is more clean.

In the `add_data()` / `add_font()` example, if writing the data with row
names

While it is possible to construct dimensions from decreasing rows and
columns, the output will always order the rows top to bottom. So
`wb_dims(rows = 3:1, cols = 3:1)` will not result in `"C3:A1"` and if
passed to functions, it will return the same as `"C1:A3"`.

## Using `wb_dims()` without an `x` object

- `rows` / `cols` (if you want to specify a single one, use `from_row` /
  `from_col`)

- `from_row` / `from_col` the starting position of the `dims` (similar
  to `start_row` / `start_col`, but with a clearer name.)

## Using `wb_dims()` with an `x` object

`wb_dims()` with an object has 8 use-cases (they work with any position
values of `from_row` / `from_col`), `from_col/from_row` correspond to
the coordinates at the top left of `x` including column and row names if
present.

These use cases are provided without `from_row / from_col`, but they
work also with `from_row / from_col`.

1.  provide the full grid with `wb_dims(x = mtcars)`

2.  provide the data grid `wb_dims(x = mtcars, select = "data")`

3.  provide the `dims` of column names
    `wb_dims(x = mtcars, select = "col_names")`

4.  provide the `dims` of row names
    `wb_dims(x = mtcars, row_names = TRUE, select = "row_names")`

5.  provide the `dims` of a row span `wb_dims(x = mtcars, rows = 1:10)`
    selects the first 10 data rows of `mtcars` (ignoring column names)

6.  provide the `dims` of the data in a column span
    `wb_dims(x = mtcars, cols = 1:5)` select the data first 5 columns of
    `mtcars`

7.  provide a column span (including column names)
    `wb_dims(x = mtcars, cols = 4:7, select = "x")` select the data
    columns 4, 5, 6, 7 of `mtcars` + column names

8.  provide the position of a single column by name
    `wb_dims(x = mtcars, cols = "mpg")`.

9.  provide a row span with a column.
    `wb_dims(x = mtcars, cols = "mpg", rows = 5:22)`

To reuse, a good trick is to create a wrapper function, so that styling
can be performed seamlessly.

    wb_dims_cars <- function(...) {
      wb_dims(x = mtcars, from_row = 2, from_col = "B", ...)
    }
    # using this function
    wb_dims_cars()                     # full grid (data + column names)
    wb_dims_cars(select = "data")      # data only
    wb_dims_cars(select = "col_names") # select column names
    wb_dims_cars(cols = "vs")          # select the `vs` column

It can be very useful to apply many rounds of styling sequentially.

## Examples

``` r
# Provide coordinates
wb_dims(1, 4)
#> [1] "D1"
wb_dims(rows = 1, cols = 4)
#> [1] "D1"
wb_dims(from_row = 4)
#> [1] "A4"
wb_dims(from_col = 2)
#> [1] "B1"
wb_dims(from_col = "B")
#> [1] "B1"
wb_dims(1:4, 6:9, from_row = 5)
#> [1] "F5:I8"
# Provide vectors
wb_dims(1:10, c("A", "B", "C"))
#> [1] "A1:C10"
wb_dims(rows = 1:10, cols = 1:10)
#> [1] "A1:J10"
# provide `from_col` / `from_row`
wb_dims(rows = 1:10, cols = c("A", "B", "C"), from_row = 2)
#> [1] "A2:C11"
wb_dims(rows = 1:10, cols = 1:10, from_col = 2)
#> [1] "B1:K10"
wb_dims(rows = 1:10, cols = 1:10, from_dims = "B1")
#> [1] "B1:K10"
# or objects
wb_dims(x = mtcars, col_names = TRUE)
#> [1] "A1:K33"

# select all data
wb_dims(x = mtcars, select = "data")
#> [1] "A2:K33"

# column names of an object (with the special select = "col_names")
wb_dims(x = mtcars, select = "col_names")
#> [1] "A1:K1"


# dims of the column names of an object
wb_dims(x = mtcars, select = "col_names", col_names = TRUE)
#> [1] "A1:K1"

## add formatting to `mtcars` using `wb_dims()`----
wb <- wb_workbook()
wb$add_worksheet("test wb_dims() with an object")
dims_mtcars_and_col_names <- wb_dims(x = mtcars)
wb$add_data(x = mtcars, dims = dims_mtcars_and_col_names)

# Put the font as Arial for the data
dims_mtcars_data <- wb_dims(x = mtcars, select = "data")
wb$add_font(dims = dims_mtcars_data, name = "Arial")

# Style col names as bold using the special `select = "col_names"` with `x` provided.
dims_column_names <- wb_dims(x = mtcars, select = "col_names")
wb$add_font(dims = dims_column_names, bold = TRUE, size = 13)

# Finally, to add styling to column "cyl" (the 4th column) (only the data)
# there are many options, but here is the preferred one
# if you know the column index, wb_dims(x = mtcars, cols = 4) also works.
dims_cyl <- wb_dims(x = mtcars, cols = "cyl")
wb$add_fill(dims = dims_cyl, color = wb_color("pink"))

# Mark a full column as important(with the column name too)
wb_dims_vs <- wb_dims(x = mtcars, cols = "vs", select = "x")
wb$add_fill(dims = wb_dims_vs, fill = wb_color("yellow"))
#> Warning: unused arguments (fill)
wb$add_conditional_formatting(dims = wb_dims(x = mtcars, cols = "mpg"), type = "dataBar")
# wb_open(wb)

# fix relative ranges
wb_dims(x = mtcars) # equal to none: A1:K33
#> [1] "A1:K33"
wb_dims(x = mtcars, fix = "all") # $A$1:$K$33
#> [1] "$A$1:$K$33"
wb_dims(x = mtcars, fix = "row") # A$1:K$33
#> [1] "A$1:K$33"
wb_dims(x = mtcars, fix = "col") # $A1:$K33
#> [1] "$A1:$K33"
wb_dims(x = mtcars, fix = c("col", "row")) # $A1:K$33
#> [1] "$A1:K$33"
```
