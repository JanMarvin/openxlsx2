# Modify column widths of a worksheet

Remove / set worksheet column widths to specified width or "auto".

## Usage

``` r
wb_set_col_widths(
  wb,
  sheet = current_sheet(),
  cols,
  widths = 8.43,
  hidden = FALSE
)

wb_remove_col_widths(wb, sheet = current_sheet(), cols)
```

## Arguments

- wb:

  A `wbWorkbook` object.

- sheet:

  A name or index of a worksheet, a vector in the case of `remove_`

- cols:

  Indices of cols to set/remove column widths.

- widths:

  Width to set `cols` to specified column width or `"auto"` for
  automatic sizing. `widths` is recycled to the length of `cols`.
  openxlsx2 sets the default width is 8.43, as this is the standard in
  some spreadsheet software. See **Details** for general information on
  column widths.

- hidden:

  Logical vector recycled to the length of `cols`. If `TRUE`, the
  columns are hidden.

## Details

The global minimum and maximum column width for "auto" columns are
controlled by:

- `options("openxlsx2.minWidth" = 3)`

- `options("openxlsx2.maxWidth" = 250)` (the maximum width allowed in
  OOXML)

Automatic column width calculation is a heuristic that may not be
accurate in all scenarios. Known limitations include issues with wrapped
text, merged cells, and font styles with variable kerning. The
underlying logic primarily assumes a monospace font and provides limited
support for specific number formats. As a safeguard to avoid very narrow
columns, widths calculated below the `openxlsx2.minWidth` (or if unset,
below 4) threshold are slightly increased.

Be aware that calculating widths can be computationally slow for large
worksheets. Additionally, the `hidden` parameter is linked with settings
in
[`wb_group_cols()`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
so changing one will update the other. Because default column widths are
influenced by the specific spreadsheet software, operating system, and
DPI settings, even providing specific values for `widths` does not
guarantee perfectly consistent output across all environments.

For automatic text wrapping of columns use [wb_add_cell_style(wrap_text
=
TRUE)](https://janmarvin.github.io/openxlsx2/reference/wb_add_cell_style.md)

## See also

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md),
[`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_chartsheet.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md),
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

Other worksheet content functions:
[`filter-wb`](https://janmarvin.github.io/openxlsx2/reference/filter-wb.md),
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
## Create a new workbook
wb <- wb_workbook()

## Add a worksheet
wb$add_worksheet("Sheet 1")

## set col widths
wb$set_col_widths(cols = c(1, 4, 6, 7, 9), widths = c(16, 15, 12, 18, 33))

## auto columns
wb$add_worksheet("Sheet 2")
wb$add_data(sheet = 2, x = iris)
wb$set_col_widths(sheet = 2, cols = 1:5, widths = "auto")

## removing column widths
## Create a new workbook
wb <- wb_load(file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))

## remove column widths in columns 1 to 20
wb_remove_col_widths(wb, 1, cols = 1:20)
```
