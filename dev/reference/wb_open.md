# Preview a workbook in spreadsheet software

`wb_open()` provides a convenient interface to immediately view the
contents of a `wbWorkbook` object within a spreadsheet application. This
function serves as a high-level wrapper for
[`xl_open()`](https://janmarvin.github.io/openxlsx2/dev/reference/xl_open.md),
allowing users to inspect the results of programmatic workbook
construction without explicitly managing file paths.

## Usage

``` r
wb_open(wb, interactive = NA, flush = FALSE)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object to be previewed.

- interactive:

  Logical; determines if the file should be opened. When `NA` (the
  default), it inherits the value from
  [`base::interactive()`](https://rdrr.io/r/base/interactive.html). If
  `FALSE`, a warning is issued and the file is not opened.

- flush:

  Logical; if `TRUE`, the `flush` argument is passed to the internal
  save call. This controls the XML processing method used when writing
  the temporary file. For a detailed discussion on the performance and
  memory implications of this parameter, see
  [`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md).

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
object, invisibly.

## Details

The function operates by creating a temporary copy of the workbook on
the local file system and subsequently invoking the system's default
handler or a specified spreadsheet application. For users utilizing the
R6 interface, `wb$open()` is available as a shorter alias for this
function.

## See also

[`xl_open()`](https://janmarvin.github.io/openxlsx2/dev/reference/xl_open.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md)
