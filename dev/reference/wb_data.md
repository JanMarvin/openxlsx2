# Add the `wb_data` attribute to a data frame in a worksheet

provide wb_data object as mschart input

## Usage

``` r
wb_data(
  wb,
  sheet = current_sheet(),
  dims,
  header_rows = 1L,
  used_range = TRUE,
  ...
)

# S3 method for class 'wb_data'
`[`(
  x,
  i,
  j,
  drop = !((missing(j) && length(i) > 1) || (!missing(i) && length(j) > 1))
)
```

## Arguments

- wb:

  a workbook

- sheet:

  a sheet in the workbook either name or index

- dims:

  the dimensions

- header_rows:

  number of rows at the top of `dims` that form the column header.

- used_range:

  if `TRUE`, full column dims like `"A:B"` and full row dims like
  `"1:2"` read only the used range of the worksheet plus one blank row
  or column. The requested dims are kept in the attribute
  `dims_requested`, which pivot tables use as their source. If `FALSE`,
  all rows or columns are read.

- ...:

  additional arguments for
  [`wb_to_df()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_to_df.md).
  Be aware that not every argument is valid.

- x:

  x

- i:

  i

- j:

  j

- drop:

  drop

## Value

A data frame of class `wb_data`.

## See also

[`wb_to_df()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_to_df.md)
[`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_mschart.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md)

## Examples

``` r
 wb <- wb_workbook()
 wb <- wb_add_worksheet(wb)
 wb <- wb_add_data(wb, x = mtcars, dims = "B2")

 wb_data(wb, 1, dims = "B2:E6")
#>    mpg cyl disp  hp
#> 3 21.0   6  160 110
#> 4 21.0   6  160 110
#> 5 22.8   4  108  93
#> 6 21.4   6  258 110

 # a header of two rows
 wb <- wb_workbook()
 wb <- wb_add_worksheet(wb)
 wb <- wb_add_data(
   wb,
   x = matrix(c("A", NA, "B", NA, "x", "y", "x", "y", 1, 2, 3, 4), nrow = 3, byrow = TRUE),
   col_names = FALSE
 )
 wb_data(wb, 1, dims = "A1:D3", header_rows = 2)
#>   A x A y B x B y
#> 3   1   2   3   4

 # full columns are limited to the used range
 wb_data(wb, 1, dims = "A:D")
#>      A <NA>    B <NA>
#> 2    x    y    x    y
#> 3    1    2    3    4
#> 4 <NA> <NA> <NA> <NA>
```
