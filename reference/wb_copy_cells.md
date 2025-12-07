# Copy cells around within a worksheet

Copy cells around within a worksheet

## Usage

``` r
wb_copy_cells(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  data,
  as_value = FALSE,
  as_ref = FALSE,
  transpose = FALSE,
  ...
)
```

## Arguments

- wb:

  A workbook

- sheet:

  a worksheet

- dims:

  A cell where to place the copy

- data:

  A
  [`wb_data`](https://janmarvin.github.io/openxlsx2/reference/wb_data.md)
  object containing cells to copy

- as_value:

  Should a copy of the value be written?

- as_ref:

  Should references to the cell be written?

- transpose:

  Should the data be written transposed?

- ...:

  additional arguments passed to add_data() if used with `as_value`

## Value

the `wbWorkbook` invisibly

## See also

[`wb_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_data.md)

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
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

## Examples

``` r
wb <- wb_workbook()$
add_worksheet()$
  add_data(x = mtcars)$
  add_fill(dims = "A1:F1", color = wb_color("yellow"))

dat <- wb_data(wb, dims = "A1:D4", col_names = FALSE)
# 1:1 copy to M2
wb$
  clone_worksheet(old = 1, new = "Clone1")$
  copy_cells(data = dat, dims = "M2")
```
