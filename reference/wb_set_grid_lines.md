# Modify grid lines visibility in a worksheet

Set worksheet grid lines to show or hide. You can also add / remove grid
lines when creating a worksheet with
[`wb_add_worksheet(grid_lines = FALSE)`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md)

## Usage

``` r
wb_set_grid_lines(wb, sheet = current_sheet(), show = FALSE, print = show)

wb_grid_lines(wb, sheet = current_sheet(), show = FALSE, print = show)
```

## Arguments

- wb:

  A workbook object

- sheet:

  A name or index of a worksheet

- show:

  A logical. If `FALSE`, grid lines are hidden.

- print:

  A logical. If `FALSE`, grid lines are not printed.

## Examples

``` r
wb <- wb_workbook()$add_worksheet()$add_worksheet()
wb$get_sheet_names() ## list worksheets in workbook
#>   Sheet 1   Sheet 2 
#> "Sheet 1" "Sheet 2" 
wb$set_grid_lines(1, show = FALSE)
wb$set_grid_lines("Sheet 2", show = FALSE)
```
