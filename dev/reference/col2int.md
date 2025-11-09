# Convert spreadsheet column to integer

Converts a spreadsheet column in `A1` notation to an integer. `"A"` is
`1`, `"B"` is `2`, ..., `"Z"` is `26` and `"AA"` is `27`.

## Usage

``` r
col2int(x)
```

## Arguments

- x:

  A character vector

## Value

An integer column label (or `NULL` if `x` is `NULL`)

## Examples

``` r
col2int(LETTERS)
#>  [1]  1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25
#> [26] 26
```
