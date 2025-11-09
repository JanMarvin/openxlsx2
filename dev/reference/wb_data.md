# Add the `wb_data` attribute to a data frame in a worksheet

provide wb_data object as mschart input

## Usage

``` r
wb_data(wb, sheet = current_sheet(), dims, ...)

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
```
