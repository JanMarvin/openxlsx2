# Open a file or workbook object in spreadsheet software

`xl_open()` is a portable utility designed to open spreadsheet files
(such as .xlsx or .xls) or
[wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
objects using the appropriate application based on the host operating
system. It handles the nuances of background execution to ensure the R
interpreter remains unblocked.

## Usage

``` r
xl_open(x, interactive = NA, flush = FALSE)

# S3 method for class 'wbWorkbook'
xl_open(x, interactive = NA, flush = FALSE)

# Default S3 method
xl_open(x, interactive = NA, flush = FALSE)
```

## Arguments

- x:

  A character string specifying the path to a spreadsheet file or a
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- interactive:

  Logical; if `FALSE`, the function will not attempt to launch the
  application. Defaults to the result of
  [`base::interactive()`](https://rdrr.io/r/base/interactive.html).

- flush:

  Logical; if `TRUE`, the workbook is written to the temporary location
  using the stream-based XML parser. See
  [`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md)
  for details.

## Details

The method for identifying and launching the software varies by
platform:

**Windows** utilizes `shell.exec()` to trigger the file association
registered with the operating system.

**macOS** utilizes the system `open` command, which respects default
application handlers for the file type. Users can override the default
by setting `options("openxlsx2.excelApp")`.

**Linux** attempts to locate common spreadsheet utilities in the system
path, including LibreOffice (`soffice`), Gnumeric (`gnumeric`), Calligra
Sheets (`calligrasheets`), and ONLYOFFICE (`onlyoffice-desktopeditors`).
If multiple applications are found during an interactive session, a menu
is presented to the user to define their preference, which is then
stored in `options("openxlsx2.excelApp")`.

For `wbWorkbook` objects, the function automatically clones the
workbook, detects the presence of macros (VBA) to determine the
appropriate temporary file extension, and saves the content to a
temporary location before opening.

## Examples

``` r
# \donttest{
if (interactive()) {
  xlsx_file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
  xl_open(xlsx_file)

  # (not yet saved) Workbook example
  wb <- wb_workbook()
  x <- mtcars[1:6, ]
  wb$add_worksheet("Cars")
  wb$add_data("Cars", x, start_col = 2, start_row = 3, row_names = TRUE)
  xl_open(wb)
}
# }
```
