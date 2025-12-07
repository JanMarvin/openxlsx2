# Order worksheets in a workbook

Get/set order of worksheets in a Workbook object

## Usage

``` r
wb_get_order(wb)

wb_set_order(wb, sheets)
```

## Arguments

- wb:

  A `wbWorkbook` object

- sheets:

  Sheet order

## Details

This function does not reorder the worksheets within the workbook
object, it simply shuffles the order when writing to file.

## Examples

``` r
## setup a workbook with 3 worksheets
wb <- wb_workbook()
wb$add_worksheet("Sheet 1", grid_lines = FALSE)
wb$add_data_table(x = iris)

wb$add_worksheet("mtcars (Sheet 2)", grid_lines = FALSE)
wb$add_data(x = mtcars)

wb$add_worksheet("Sheet 3", grid_lines = FALSE)
wb$add_data(x = Formaldehyde)

wb_get_order(wb)
#> [1] 1 2 3
wb$get_sheet_na
#> NULL
wb$set_order(c(1, 3, 2)) # switch position of sheets 2 & 3
wb$add_data(2, 'This is still the "mtcars" worksheet', start_col = 15)
wb_get_order(wb)
#> [1] 1 3 2
wb$get_sheet_names() ## ordering within workbook is not changed
#>            Sheet 1            Sheet 3   mtcars (Sheet 2) 
#>          "Sheet 1"          "Sheet 3" "mtcars (Sheet 2)" 
wb$set_order(3:1)
```
