# Add a chart XML to a worksheet

Add a chart XML to a worksheet

## Usage

``` r
wb_add_chart_xml(
  wb,
  sheet = current_sheet(),
  dims = NULL,
  xml,
  col_offset = 0,
  row_offset = 0,
  ...
)
```

## Arguments

- wb:

  a workbook

- sheet:

  the sheet on which the graph will appear

- dims:

  the dimensions where the sheet will appear

- xml:

  chart xml

- col_offset, row_offset:

  positioning

- ...:

  additional arguments

## See also

[`wb_add_drawing()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_drawing.md)
[`wb_add_image()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_image.md)
[`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_mschart.md)
[`wb_add_plot()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_plot.md)
