# Add/remove column filters in a worksheet

Add or remove spreadsheet column filters to a worksheet

## Usage

``` r
wb_add_filter(wb, sheet = current_sheet(), rows, cols)

wb_remove_filter(wb, sheet = current_sheet())
```

## Arguments

- wb:

  A workbook object

- sheet:

  A worksheet name or index. In `wb_remove_filter()`, you may supply a
  vector of worksheets.

- rows:

  A row number.

- cols:

  columns to add filter to.

## Details

Adds filters to worksheet columns, same as `with_filter = TRUE` in
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md)
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md)
automatically adds filters to first row of a table.

NOTE Can only have a single filter per worksheet unless using tables.

## See also

[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md)

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
[`named_region-wb`](https://janmarvin.github.io/openxlsx2/reference/named_region-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md),
[`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_conditional_formatting.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md)

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")
wb$add_worksheet("Sheet 2")
wb$add_worksheet("Sheet 3")

wb$add_data(1, iris)
wb$add_filter(1, row = 1, cols = seq_along(iris))

## Equivalently
wb$add_data(2, x = iris, with_filter = TRUE)

## Similarly
wb$add_data_table(3, iris)
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")
wb$add_worksheet("Sheet 2")
wb$add_worksheet("Sheet 3")

wb$add_data(1, iris)
wb_add_filter(wb, 1, row = 1, cols = seq_along(iris))

## Equivalently
wb$add_data(2, x = iris, with_filter = TRUE)

## Similarly
wb$add_data_table(3, iris)

## remove filters
wb_remove_filter(wb, 1:2) ## remove filters
wb_remove_filter(wb, 3) ## Does not affect tables!
```
