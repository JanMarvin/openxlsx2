# List tables in a worksheet

List tables in a worksheet

## Usage

``` r
wb_get_tables(wb, sheet = current_sheet())
```

## Arguments

- wb:

  A workbook object

- sheet:

  A name or index of a worksheet

## Value

A character vector of table names on the specified sheet

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet(sheet = "Sheet 1")
wb$add_data_table(x = iris)
wb$add_data_table(x = mtcars, table_name = "mtcars", start_col = 10)

wb$get_tables(sheet = "Sheet 1")
#>   tab_name tab_ref
#> 1   Table1 A1:E151
#> 2   mtcars  J1:T33
```
