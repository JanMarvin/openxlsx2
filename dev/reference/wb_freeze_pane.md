# Freeze panes of a worksheet

The `wb_freeze_pane()` function locks a specific area of a worksheet to
keep rows or columns visible while scrolling through other parts of the
data. This is achieved by defining a "split" point, where all content
above or to the left of the designated active region remains fixed.

## Usage

``` r
wb_freeze_pane(
  wb,
  sheet = current_sheet(),
  first_active_row = NULL,
  first_active_col = NULL,
  first_row = FALSE,
  first_col = FALSE,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet to modify. Defaults to the current
  sheet.

- first_active_row:

  The index of the first row that should remain scrollable. Rows above
  this will be frozen.

- first_active_col:

  The index or character label of the first column that should remain
  scrollable. Columns to the left will be frozen.

- first_row:

  Logical; if `TRUE`, freezes the first row of the worksheet.

- first_col:

  Logical; if `TRUE`, freezes the first column of the worksheet.

- ...:

  Additional arguments for internal case standardization.

## Details

The function operates by calculating `xSplit` and `ySplit` values based
on the provided active region coordinates. The `first_active_row` and
`first_active_col` parameters define the first cell that remains
scrollable; consequently, the frozen area consists of all rows and
columns preceding these indices.

For common use cases, the `first_row` and `first_col` logical flags
provide optimized shortcuts. Enabling `first_row` locks the top row
(equivalent to setting the active region at row 2), while `first_col`
locks the leftmost column (equivalent to setting the active region at
column 2). If both are enabled, the function automatically freezes the
intersection at cell "B2".

The internal logic translates these coordinates into a `<pane />` XML
node, which specifies the `topLeftCell` of the scrollable region and
assigns the `activePane` (e.g., "bottomLeft", "topRight", or
"bottomRight") to ensure correct cursor behavior within the spreadsheet
software.

## Notes

- If `first_active_row` and `first_active_col` are both set to 1, or if
  all arguments are omitted, the function returns the workbook unchanged
  as there is no region to freeze.

- This function overwrites any existing pane configuration for the
  specified worksheet.

## See also

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chartsheet.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_worksheet.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_copy_cells.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_workbook.md)

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`filter-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/filter-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`named_region-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/named_region-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_conditional_formatting.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_thread.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md)

## Examples

``` r
wb <- wb_workbook()
## Add some worksheets
wb$add_worksheet("Sheet 1")
wb$add_worksheet("Sheet 2")
wb$add_worksheet("Sheet 3")
wb$add_worksheet("Sheet 4")

## Freeze Panes
wb$freeze_pane("Sheet 1", first_active_row = 5, first_active_col = 3)
wb$freeze_pane("Sheet 2", first_col = TRUE) ## shortcut to first_active_col = 2
wb$freeze_pane(3, first_row = TRUE) ## shortcut to first_active_row = 2
wb$freeze_pane(4, first_active_row = 1, first_active_col = "D")
```
