# Add a page break to a worksheet

Insert page breaks into a worksheet

## Usage

``` r
wb_add_page_break(wb, sheet = current_sheet(), row = NULL, col = NULL)
```

## Arguments

- wb:

  A workbook object

- sheet:

  A name or index of a worksheet

- row, col:

  Either a row number of column number. One must be `NULL`

## See also

[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md)

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")
wb$add_data(sheet = 1, x = iris)

wb$add_page_break(sheet = 1, row = 10)
wb$add_page_break(sheet = 1, row = 20)
wb$add_page_break(sheet = 1, col = 2)
```
