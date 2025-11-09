# Add a formula to a cell range in a worksheet

This function can be used to add a formula to a worksheet. In
`wb_add_formula()`, you can provide the formula as a character vector.

## Usage

``` r
wb_add_formula(
  wb,
  sheet = current_sheet(),
  x,
  dims = wb_dims(start_row, start_col),
  start_col = 1,
  start_row = 1,
  array = FALSE,
  cm = FALSE,
  apply_cell_style = TRUE,
  remove_cell_style = FALSE,
  enforce = FALSE,
  shared = FALSE,
  name = NULL,
  ...
)
```

## Arguments

- wb:

  A Workbook object containing a worksheet.

- sheet:

  The worksheet to write to. (either as index or name)

- x:

  A formula as character vector.

- dims:

  Spreadsheet dimensions that will determine where `x` spans: "A1",
  "A1:B2", "A:B"

- start_col:

  A vector specifying the starting column to write to.

- start_row:

  A vector specifying the starting row to write to.

- array:

  A bool if the function written is of type array

- cm:

  A special kind of array function that hides the curly braces in the
  cell. Add this, if you see "@" inserted into your formulas.

- apply_cell_style:

  Should we write cell styles to the workbook?

- remove_cell_style:

  Should we keep the cell style?

- enforce:

  enforce dims

- shared:

  shared formula

- name:

  The name of a named region if specified.

- ...:

  additional arguments

## Value

The workbook, invisibly.

## Details

Currently, the local translations of formulas are not supported. Only
the English functions work.

The examples below show a small list of possible formulas:

- SUM(B2:B4)

- AVERAGE(B2:B4)

- MIN(B2:B4)

- MAX(B2:B4)

- ...

It is possible to pass vectors to `x`. If `x` is an array formula, it
will take `dims` as a reference. For some formulas, the result will span
multiple cells (see the `MMULT()` example below). For this type of
formula, the output range must be known a priori and passed to `dims`,
otherwise only the value of the first cell will be returned. This type
of formula, whose result extends over several cells, is only possible
with single strings. If a vector is passed, it is only possible to
return individual cells.

Custom functions can be registered as lambda functions in the workbook.
For this you take the function you want to add `"LAMBDA(x, y, x + y)"`
and escape it as follows. `LAMBDA()` is a future function and needs a
prefix `_xlfn`. The arguments need a prefix `_xlpm.`. So the full
function looks like this:
`"_xlfn.LAMBDA(_xlpm.x, _xlpm.y, _xlpm.x + _xlpm.y)"`. These custom
formulas are accessible via the named region manager and can be removed
with
[`wb_remove_named_region()`](https://janmarvin.github.io/openxlsx2/dev/reference/named_region-wb.md).
Contrary to other formulas, custom formulas must be registered with the
workbook before they can be used (see the example below).

If a function that normally works in spreadsheet software does not
behave as expected when written using `wb_add_formula()`, e.g., if
spurious `@` symbols appear in the formula, it is likely that the
formula is either an array formula or requires a future function prefix.
In modern spreadsheet software, it is no longer straightforward to
detect whether a formula is an array formula, since this hidden in cell
metadata (cm). Therefore, a formula like `SUM(1+(A1:A2))` will not be
displayed as `{SUM(1+(A1:A2))}`.

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
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`named_region-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/named_region-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_conditional_formatting.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_thread.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md)

## Examples

``` r
wb <- wb_workbook()$add_worksheet()
wb$add_data(dims = wb_dims(rows = 1, cols = 1:3), x = c(4, 5, 8))

# calculate the sum of elements.
wb$add_formula(dims = "D1", x = "SUM(A1:C1)")

# array formula with result spanning over multiple cells
mm <- matrix(1:4, 2, 2)

wb$add_worksheet()$
 add_data(x = mm, dims = "A1:B2", col_names = FALSE)$
 add_data(x = mm, dims = "A4:B5", col_names = FALSE)$
 add_formula(x = "MMULT(A1:B2, A4:B5)", dims = "A7:B8", array = TRUE)

# add shared formula
wb$add_worksheet()$
 add_data(x = matrix(1:25, ncol = 5, nrow = 5))$
 add_formula(x = "SUM($A2:A2)", dims = "A8:E12", shared = TRUE)

# add a custom formula, first define it, then use it
wb$add_formula(x = c(YESTERDAY = "_xlfn.LAMBDA(TODAY() - 1)"))
#> formula registered to the workbook
wb$add_formula(x = "=YESTERDAY()", dims = "A1", cm = TRUE)
#> Warning: modifications with cm formulas are experimental. use at own risk
```
