# Merge cells within a worksheet

Worksheet cell merging

## Usage

``` r
wb_merge_cells(
  wb,
  sheet = current_sheet(),
  dims = NULL,
  solve = FALSE,
  direction = NULL,
  ...
)

wb_unmerge_cells(wb, sheet = current_sheet(), dims = NULL, ...)
```

## Arguments

- wb:

  A Workbook object

- sheet:

  A name or index of a worksheet

- dims:

  worksheet cells

- solve:

  logical if intersecting merges should be solved

- direction:

  direction in which to split the cell merging. Allows "row" or "col"

- ...:

  additional arguments

## Details

If using the deprecated arguments `rows` and `cols` with a merged region
must be rectangular, only min and max of `cols` and `rows` are used.

## See also

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
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
[`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
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
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md)

## Examples

``` r
# Create a new workbook
wb <- wb_workbook()$add_worksheet()

# Merge cells: Row 2 column C to F (3:6)
wb <- wb_merge_cells(wb, dims = "C3:F6")

# Merge cells:Rows 10 to 20 columns A to J (1:10)
wb <- wb_merge_cells(wb, dims = wb_dims(rows = 10:20, cols = 1:10))

wb$add_worksheet()

## Intersecting merges
wb <- wb_merge_cells(wb, dims = wb_dims(cols = 1:10, rows = 1))
wb <- wb_merge_cells(wb, dims = wb_dims(cols = 5:10, rows = 2))
wb <- wb_merge_cells(wb, dims = wb_dims(cols = 1:10, rows = 12))
try(wb_merge_cells(wb, dims = "A1:A10"))
#> Error : Merge intersects with existing merged cells: 
#>      A1:J1.
#> Remove existing merge first.

## remove merged cells
# removes any intersecting merges
wb <- wb_unmerge_cells(wb, dims = wb_dims(cols = 1, rows = 1))
wb <- wb_merge_cells(wb, dims = "A1:A10")

# or let us decide how to solve this
wb <- wb_merge_cells(wb, dims = "A1:A10", solve = TRUE)
```
