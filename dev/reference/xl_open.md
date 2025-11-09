# Open an xlsx file or a `wbWorkbook` object

This function tries to open a Microsoft Excel (xls/xlsx) file or, an
[wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
with the proper application, in a portable manner.

On Windows it uses `base::shell.exec()` (Windows only function) to
determine the appropriate program.

On Mac, (c) it uses system default handlers, given the file type.

On Linux, it searches (via `which`) for available xls/xlsx reader
applications (unless `options('openxlsx2.excelApp')` is set to the app
bin path), and if it finds anything, sets
`options('openxlsx2.excelApp')` to the program chosen by the user via a
menu (if many are present, otherwise it will set the only available).
Currently searched for apps are Libreoffice/Openoffice (`soffice` bin),
Gnumeric (`gnumeric`), Calligra Sheets (`calligrasheets`) and ONLYOFFICE
(`onlyoffice-desktopeditors`).

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

  A path to a spreadsheet file or wbWorkbook object. This can be any
  file type that can be opened in the corresponding software.

- interactive:

  If `FALSE` will throw a warning and not open the path. This can be
  manually set to `TRUE`, otherwise when `NA` (default) uses the value
  returned from
  [`base::interactive()`](https://rdrr.io/r/base/interactive.html)

- flush:

  If `TRUE` the `flush` argument of
  [`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md)
  will be used to create the output file. Applies only to workbooks.

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
