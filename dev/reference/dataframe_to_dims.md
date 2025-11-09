# Create dimensions from dataframe

Use
[`wb_dims()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_dims.md)

## Usage

``` r
dataframe_to_dims(df, dim_break = FALSE)
```

## Arguments

- df:

  dataframe with spreadsheet columns and rows

- dim_break:

  split the dims?

## Examples

``` r
 df <- dims_to_dataframe("A1:D5;F1:F6;D8", fill = TRUE)
 dataframe_to_dims(df)
#> [1] "A1:F8"
```
