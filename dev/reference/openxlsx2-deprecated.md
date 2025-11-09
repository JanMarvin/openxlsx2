# Deprecated functions in package *openxlsx2*

These functions are provided for compatibility with older versions of
`openxlsx2`, and may be defunct as soon as the next release. This guide
helps you update your code to the latest standards.

As of openxlsx2 v1.0, API change should be minimal.

## Internal functions

These functions are used internally by openxlsx2. It is no longer
advertised to use them in scripts. They originate from openxlsx, but do
not fit openxlsx2's API.

You should be able to modify

- [`delete_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/delete_data.md)
  -\>
  [`wb_clean_sheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_clean_sheet.md)

- [`write_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/write_data.md)
  -\>
  [`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md)

- [`write_datatable()`](https://janmarvin.github.io/openxlsx2/dev/reference/write_datatable.md)
  -\>
  [`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md)

- [`write_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/comment_internal.md)
  -\>
  [`wb_add_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_comment.md)

- [`remove_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/comment_internal.md)
  -\>
  [`wb_remove_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_comment.md)

- [`write_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/write_formula.md)
  -\>
  [`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md)

You should be able to change those with minimal changes

## Deprecated functions

First of all, you can set an option that will add warnings when using
deprecated functions.

    options("openxlsx2.soon_deprecated" = TRUE)

## Argument changes

For consistency, arguments were renamed to snake_case for the 0.8
release. It is now recommended to use `dims` (the cell range) in favor
of `row`, `col`, `start_row`, `start_col`

See
[`wb_dims()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_dims.md)
as it provides many options on how to provide cell range

## Functions with a new name

These functions were renamed for consistency.

- [`convertToExcelDate()`](https://janmarvin.github.io/openxlsx2/dev/reference/convertToExcelDate.md)
  -\>
  [`convert_to_excel_date()`](https://janmarvin.github.io/openxlsx2/dev/reference/convert_to_excel_date.md)

- [`wb_grid_lines()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_grid_lines.md)
  -\>
  [`wb_set_grid_lines()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_grid_lines.md)

- [`create_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_comment.md)
  -\>
  [`wb_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_comment.md)

## Deprecated usage

- [`wb_get_named_regions()`](https://janmarvin.github.io/openxlsx2/dev/reference/named_region-wb.md)
  will no longer allow providing a file.

    ## Before
    wb_get_named_regions(file)

    ## Now
    wb <- wb_load(file)
    wb_get_named_regions(wb)
    # also possible
    wb_load(file)$get_named_regions()`

## See also

[.Deprecated](https://rdrr.io/r/base/Deprecated.html)
