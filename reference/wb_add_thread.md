# Add threaded comments to a cell in a worksheet

These functions allow adding thread comments to spreadsheets. This is
not yet supported by all spreadsheet software. A threaded comment must
be tied to a person created by
[`wb_add_person()`](https://janmarvin.github.io/openxlsx2/reference/person-wb.md).

## Usage

``` r
wb_add_thread(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  comment = NULL,
  person_id,
  reply = FALSE,
  resolve = FALSE
)

wb_get_thread(wb, sheet = current_sheet(), dims = NULL)
```

## Arguments

- wb:

  A workbook

- sheet:

  A worksheet

- dims:

  A cell

- comment:

  The text to add, a character vector.

- person_id:

  the person Id this should be added. The default is
  `getOption("openxlsx2.thread_id")` if set.

- reply:

  Is the comment a reply? (default `FALSE`)

- resolve:

  Should the comment be resolved? (default `FALSE`)

## Details

If a threaded comment is added, it needs a person attached to it. The
default is to create a person with provider id `"None"`. Other providers
are possible with specific values for `id` and `user_id`. If you require
the following, create a workbook via spreadsheet software load it and
get the values with
[`wb_get_person()`](https://janmarvin.github.io/openxlsx2/reference/person-wb.md)

## See also

[`wb_add_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_comment.md)
[`person-wb`](https://janmarvin.github.io/openxlsx2/reference/person-wb.md)

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
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md)

## Examples

``` r
wb <- wb_workbook()$add_worksheet()
# Add a person to the workbook.
wb$add_person(name = "someone who likes to edit workbooks")

pid <- wb$get_person(name = "someone who likes to edit workbooks")$id

# write a comment to a thread, reply to one and solve some
wb <- wb_add_thread(wb, dims = "A1", comment = "wow it works!", person_id = pid)
wb <- wb_add_thread(wb, dims = "A2", comment = "indeed", person_id = pid, resolve = TRUE)
wb <- wb_add_thread(wb, dims = "A1", comment = "so cool", person_id = pid, reply = TRUE)
```
