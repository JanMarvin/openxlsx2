# Create dataframe from dimensions

Non consecutive decreasing dims will return an increasing data frame.

## Usage

``` r
dims_to_dataframe(dims, fill = FALSE, empty_rm = FALSE)
```

## Arguments

- dims:

  Character vector of expected dimension.

- fill:

  If `TRUE`, fills the dataframe with variables

- empty_rm:

  Logical if empty columns and rows should be included

## Examples

``` r
dims_to_dataframe("A1:B2")
#>      A    B
#> 1 <NA> <NA>
#> 2 <NA> <NA>
```
