# Save a workbook to file

Save a workbook to file

## Usage

``` r
wb_save(wb, file = NULL, overwrite = TRUE, path = NULL, flush = FALSE)
```

## Arguments

- wb:

  A `wbWorkbook` object to write to file

- file:

  A path to save the workbook to

- overwrite:

  If `FALSE`, will not overwrite when `file` already exists.

- path:

  Deprecated argument. Please use `file` in new code.

- flush:

  Experimental, streams the worksheet file to disk

## Value

the `wbWorkbook` object, invisibly

## Details

When saving a `wbWorkbook` to a file, memory usage may spike depending
on the worksheet size. This happens because the entire XML structure is
created in memory before writing to disk. The memory required depends on
worksheet size, as XML files consist of character data and include
additional overhead for validity checks.

The `flush` argument streams worksheet XML data directly to disk,
avoiding the need to build the full XML tree in memory. This reduces
memory usage but skips some XML validity checks. It also bypasses the
`pugixml` functions that `openxlsx2` uses, omitting certain preliminary
sanity checks before writing. As the name suggests, the output is simply
flushed to disk.

By default, the [`utils::zip()`](https://rdrr.io/r/utils/zip.html)
function is used to create output files. This requires a working zip
utility to be available on the system. A valid zip program must be found
either via `Sys.which("zip")` or through the `R_ZIPCMD` environment
variable.

On Windows, a suitable zip tool is typically provided by Rtools. If
`R_ZIPCMD` is not set, `openxlsx2` will automatically use the first
detected Rtools installation. If no zip utility is available, `bsdtar`
can be used as an alternative. On Windows this is shipped as `tar.exe`;
on Mac and Linux it is usually available as `bsdtar` (often requiring
installation of the `archive` package).

A further fallback—primarily for older Windows systems—is to point
`R_ZIPCMD` to `7z.exe`. This approach has not been extensively tested
and is not reliable with 7-Zip on macOS.

As an additional fallback, the `zip` package can be used. It is no
longer listed in `Imports` and must be installed separately if needed.

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
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md)

## Examples

``` r
## Create a new workbook and add a worksheet
wb <- wb_workbook("Creator of workbook")
wb$add_worksheet(sheet = "My first worksheet")

## Save workbook to working directory
# \donttest{
wb_save(wb, file = temp_xlsx(), overwrite = TRUE)

## do not use utils::zip, will try to use bsdtar
# options("openxlsx2.no_utils_zip" = TRUE)

## if the above is set as well, do not use bsdtar
# options("openxlsx2.no_bsdtar" = TRUE)

## use 7zip on Windows this works, on Mac not
# Sys.setenv("R_ZIPCMD" = "C:\Program Files\7-Zip\7z.exe")

# if the last one is left blank the fallback is zip::zip
openxlsx2::write_xlsx(x = cars, temp_xlsx())
# }
```
