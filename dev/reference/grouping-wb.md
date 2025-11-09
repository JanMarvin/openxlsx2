# Group rows and columns in a worksheet

Group a selection of rows or cols

## Usage

``` r
wb_group_cols(
  wb,
  sheet = current_sheet(),
  cols,
  collapsed = FALSE,
  levels = NULL
)

wb_ungroup_cols(wb, sheet = current_sheet(), cols)

wb_group_rows(
  wb,
  sheet = current_sheet(),
  rows,
  collapsed = FALSE,
  levels = NULL
)

wb_ungroup_rows(wb, sheet = current_sheet(), rows)
```

## Arguments

- wb:

  A `wbWorkbook` object

- sheet:

  A name or index of a worksheet

- collapsed:

  If `TRUE` the grouped columns are collapsed

- levels:

  levels

- rows, cols:

  Indices or for columns also characters of rows and columns to group

## Details

If row was previously hidden, it will now be shown. Columns can be added
using A1 notion, so `cols = 2:3` is similar to `cols = "B:C"`. It is
possible to add nested groups, so `cols = list("3" = list(1:2, 3:4)` is
also possible. Depending on the selected summary column either left or
right will be selected for grouping, this can be changed in
[`wb_set_page_setup()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_page_setup.md).

## See also

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/creators-wb.md),
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
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_workbook.md)

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`filter-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/filter-wb.md),
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
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md)

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/creators-wb.md),
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
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_workbook.md)

## Examples

``` r
# create matrix
t1 <- AirPassengers
t2 <- do.call(cbind, split(t1, cycle(t1)))
dimnames(t2) <- dimnames(.preformat.ts(t1))

wb <- wb_workbook()
wb$add_worksheet("AirPass")
wb$add_data("AirPass", t2, row_names = TRUE)

# groups will always end on/show the last row. in the example 1950, 1955, and 1960
wb <- wb_group_rows(wb, "AirPass", 2:3, collapsed = TRUE) # group years < 1950
wb <- wb_group_rows(wb, "AirPass", 4:8, collapsed = TRUE) # group years 1951-1955
wb <- wb_group_rows(wb, "AirPass", 9:13)                  # group years 1956-1960

wb <- wb_group_cols(wb, "AirPass", 2:4, collapsed = TRUE)
wb <- wb_group_cols(wb, "AirPass", 5:7, collapsed = TRUE)
wb <- wb_group_cols(wb, "AirPass", 8:10, collapsed = TRUE)
wb <- wb_group_cols(wb, "AirPass", 11:13)

### create grouping levels
grp_rows <- list(
  "1" = seq(2, 3),
  "2" = seq(4, 8),
  "3" = seq(9, 13)
)

grp_cols <- list(
  "1" = seq(2, 4),
  "2" = seq(5, 7),
  "3" = seq(8, 10),
  "4" = seq(11, 13)
)

wb <- wb_workbook()
wb$add_worksheet("AirPass")
wb$add_data("AirPass", t2, row_names = TRUE)

wb$group_cols("AirPass", cols = grp_cols)
wb$group_rows("AirPass", rows = grp_rows)
```
