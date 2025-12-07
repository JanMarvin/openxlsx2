# Add sparklines to a worksheet

Add sparklines to a worksheet

## Usage

``` r
wb_add_sparklines(wb, sheet = current_sheet(), sparklines)
```

## Arguments

- wb:

  A `wbWorkbook`

- sheet:

  sheet to add the sparklines to

- sparklines:

  sparklines object created with
  [`create_sparklines()`](https://janmarvin.github.io/openxlsx2/reference/create_sparklines.md)

## See also

[`create_sparklines()`](https://janmarvin.github.io/openxlsx2/reference/create_sparklines.md)

## Examples

``` r
 sl <- create_sparklines("Sheet 1", dims = "A3:K3", sqref = "L3")
 wb <- wb_workbook()
 wb <- wb_add_worksheet(wb)
 wb <- wb_add_data(wb, x = mtcars)
 wb <- wb_add_sparklines(wb, sparklines = sl)
```
