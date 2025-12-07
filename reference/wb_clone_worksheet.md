# Create copies of a worksheet within a workbook

Create a copy of a worksheet in the same `wbWorkbook` object.

Cloning is possible only to a limited extent. References to sheet names
in formulas, charts, pivot tables, etc. may not be updated. Some
elements like named ranges and slicers cannot be cloned yet.

Cloning from another workbook is still an experimental feature and might
not work reliably. Cloning data, media, charts and tables should work.
Slicers and pivot tables as well as everything everything relying on
dxfs styles (e.g. custom table styles and conditional formatting) is
currently not implemented. Formula references are not updated to reflect
interactions between workbooks.

## Usage

``` r
wb_clone_worksheet(wb, old = current_sheet(), new = next_sheet(), from = NULL)
```

## Arguments

- wb:

  A `wbWorkbook` object

- old:

  Name of existing worksheet to copy

- new:

  Name of the new worksheet to create

- from:

  (optional) Workbook to clone old from

## Value

The `wbWorkbook` object, invisibly.

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
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

## Examples

``` r
# Create a new workbook
wb <- wb_workbook()

# Add worksheets
wb$add_worksheet("Sheet 1")
wb$clone_worksheet("Sheet 1", new = "Sheet 2")
# Take advantage of waiver functions
wb$clone_worksheet(old = "Sheet 1")

## cloning from another workbook

# create a workbook
wb <- wb_workbook()$
add_worksheet("NOT_SUM")$
  add_data(x = head(iris))$
  add_fill(dims = "A1:B2", color = wb_color("yellow"))$
  add_border(dims = "B2:C3")

# we will clone this styled chart into another workbook
fl <- system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2")
wb_from <- wb_load(fl)

# clone styles and shared strings
wb$clone_worksheet(old = "SUM", new = "SUM", from = wb_from)
```
