# wb_add_hyperlink

Helper to add shared hyperlinks into a worksheet or remove shared
hyperlinks from a worksheet

## Usage

``` r
wb_add_hyperlink(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  target = NULL,
  tooltip = NULL,
  is_external = TRUE,
  col_names = FALSE
)

wb_remove_hyperlink(wb, sheet = current_sheet(), dims = NULL)
```

## Arguments

- wb:

  A Workbook object containing a worksheet.

- sheet:

  The worksheet to write to. (either as index or name)

- dims:

  Spreadsheet dimensions that will determine where the hyperlink
  reference spans: "A1", "A1:B2", "A:B"

- target:

  An optional target, if no target is specified, it is assumed that the
  cell already contains a reference (the cell could be a url or a
  filename)

- tooltip:

  An optional description for a variable that will be visible when
  hovering over the link text in the spreadsheet

- is_external:

  A logical indicating if the hyperlink is external (a url, a mail
  address, a file) or internal (a reference to worksheet cells)

- col_names:

  Whether or not the object contains column names. If yes the first
  column of the dimension will be ignored

## Details

There are multiple ways to add hyperlinks into a worksheet. One way is
to construct a formula with
[`create_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_hyperlink.md)
another is to assign a class `hyperlink` to a column of a data frame.
Contrary to the previous method, shared hyperlinks are not cell formulas
in the worksheet, but references in the worksheet relationship and
hyperlinks in the worksheet xml structure. These shared hyperlinks can
be reused and they are not visible to spreadsheet users as `HYPERLINK()`
formulas.

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
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`named_region-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/named_region-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_conditional_formatting.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_thread.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md)

## Examples

``` r
wb <- wb_workbook()$add_worksheet()$
  add_data(x = "openxlsx2 on CRAN")$
  add_hyperlink(target = "https://cran.r-project.org/package=openxlsx2",
                tooltip = "The canonical form to link to our CRAN page.")

wb$remove_hyperlink()
```
