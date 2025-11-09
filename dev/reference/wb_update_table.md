# Update a data table position in a worksheet

Update the position of a data table, possibly written using
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md)

## Usage

``` r
wb_update_table(wb, sheet = current_sheet(), dims = "A1", tabname)
```

## Arguments

- wb:

  A workbook

- sheet:

  A worksheet

- dims:

  Cell range used for new data table.

- tabname:

  A table name

## Details

Be aware that this function does not alter any filter. Excluding or
adding rows does not make rows appear nor will it hide them.

## Examples

``` r
wb <- wb_workbook()$add_worksheet()$add_data_table(x = mtcars)
wb$update_table(tabname = "Table1", dims = "A1:J4")
```
