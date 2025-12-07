# Preview a workbook in a spreadsheet software

You can also use the shorter `wb$open()` as a replacement. To open xlsx
files, see
[`xl_open()`](https://janmarvin.github.io/openxlsx2/reference/xl_open.md).

## Usage

``` r
wb_open(wb, interactive = NA, flush = FALSE)
```

## Arguments

- wb:

  a
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object

- interactive:

  If `FALSE` will throw a warning and not open the path. This can be
  manually set to `TRUE`, otherwise when `NA` (default) uses the value
  returned from
  [`base::interactive()`](https://rdrr.io/r/base/interactive.html)

- flush:

  if the flush option should be used
