# Convert integer to spreadsheet column

Converts an integer to a spreadsheet column in `A1` notation. `1` is
`"A"`, `2` is `"B"`, ..., `26` is `"Z"` and `27` is `"AA"`.

## Usage

``` r
int2col(x)
```

## Arguments

- x:

  A numeric vector.

## Examples

``` r
int2col(1:10)
#>  [1] "A" "B" "C" "D" "E" "F" "G" "H" "I" "J"
```
