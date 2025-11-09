# Remove a data table from a worksheet

Remove tables in a workbook using its name.

## Usage

``` r
wb_remove_tables(wb, sheet = current_sheet(), table, remove_data = TRUE)
```

## Arguments

- wb:

  A Workbook object

- sheet:

  A name or index of a worksheet

- table:

  Name of table to remove. Use
  [`wb_get_tables()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_get_tables.md)
  to view the tables present in the worksheet.

- remove_data:

  Default `TRUE`. If `FALSE`, will only remove the data table attributes
  but will keep the data in the worksheet.

## Value

The `wbWorkbook`, invisibly

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet(sheet = "Sheet 1")
wb$add_worksheet(sheet = "Sheet 2")
wb$add_data_table(sheet = "Sheet 1", x = iris, table_name = "iris")
wb$add_data_table(sheet = 1, x = mtcars, table_name = "mtcars", start_col = 10)

## delete worksheet removes table objects
wb <- wb_remove_worksheet(wb, sheet = 1)

wb$add_data_table(sheet = 1, x = iris, table_name = "iris")
wb$add_data_table(sheet = 1, x = mtcars, table_name = "mtcars", start_col = 10)

## wb_remove_tables() deletes table object and all data
wb_get_tables(wb, sheet = 1)
#>   tab_name tab_ref
#> 3     iris A1:E151
#> 4   mtcars  J1:T33
wb$remove_tables(sheet = 1, table = "iris")
wb$add_data_table(sheet = 1, x = iris, table_name = "iris")

wb_get_tables(wb, sheet = 1)
#>   tab_name tab_ref
#> 4   mtcars  J1:T33
#> 5     iris A1:E151
wb$remove_tables(sheet = 1, table = "iris")
```
