# Add drawings to a worksheet

Add drawings to a worksheet. This requires the `rvg` package.

## Usage

``` r
wb_add_drawing(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  xml,
  col_offset = 0,
  row_offset = 0,
  ...
)
```

## Arguments

- wb:

  A `wbWorkbook`

- sheet:

  A sheet in the workbook

- dims:

  The dimension where the drawing is added.

- xml:

  the drawing xml as character or file

- col_offset, row_offset:

  offsets for column and row

- ...:

  additional arguments

## See also

[`wb_add_chart_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chart_xml.md)
[`wb_add_image()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_image.md)
[`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_mschart.md)
[`wb_add_plot()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_plot.md)

## Examples

``` r
if (requireNamespace("rvg") && interactive()) {

## rvg example
require(rvg)
tmp <- tempfile(fileext = ".xml")
dml_xlsx(file =  tmp)
plot(1,1)
dev.off()

wb <- wb_workbook()$
  add_worksheet()$
  add_drawing(xml = tmp)$
  add_drawing(xml = tmp, dims = NULL)
}
#> Loading required namespace: rvg
```
