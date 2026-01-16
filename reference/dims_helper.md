# Helper functions to work with `dims`

Internal helpers to (de)construct a dims argument from/to a row and
column vector. Exported for user convenience.

## Usage

``` r
dims_to_rowcol(x, as_integer = FALSE)

validate_dims(x)

rowcol_to_dims(row, col, single = TRUE, fix = NULL)
```

## Arguments

- x:

  a dimension object "A1" or "A1:A1"

- as_integer:

  If the output should be returned as integer, (defaults to string)

- row:

  a numeric vector of rows

- col:

  a numeric or character vector of cols

- single:

  argument indicating if `rowcol_to_dims()` returns a single cell
  dimension

- fix:

  setting the type of the reference. Per default, no type is set.
  Options are `"all"`, `"row"`, and `"col"`

## Value

- A `dims` string for `_to_dim` i.e "A1:A1"

- A named list of rows and columns for `to_rowcol`

## See also

[`wb_dims()`](https://janmarvin.github.io/openxlsx2/reference/wb_dims.md)

## Examples

``` r
dims_to_rowcol("A1:J10")
#> $col
#>  [1] "A" "B" "C" "D" "E" "F" "G" "H" "I" "J"
#> 
#> $row
#>  [1] "1"  "2"  "3"  "4"  "5"  "6"  "7"  "8"  "9"  "10"
#> 
wb_dims(1:10, 1:10)
#> [1] "A1:J10"
```
