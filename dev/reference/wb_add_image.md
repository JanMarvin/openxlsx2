# Insert an image into a worksheet

Insert an image into a worksheet

## Usage

``` r
wb_add_image(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  file,
  width = 6,
  height = 3,
  row_offset = 0,
  col_offset = 0,
  units = "in",
  dpi = 300,
  address = NULL,
  ...
)
```

## Arguments

- wb:

  A workbook object

- sheet:

  A name or index of a worksheet

- dims:

  Dimensions where to plot. Default absolute anchor, single cell (eg.
  "A1") oneCellAnchor, cell range (eg. "A1:D4") twoCellAnchor

- file:

  An image file. Valid file types are:` "jpeg"`, `"png"`, `"bmp"`

- width:

  Width of figure.

- height:

  Height of figure.

- row_offset:

  offset vector for one or two cell anchor within cell (row)

- col_offset:

  offset vector for one or two cell anchor within cell (column)

- units:

  Units of width and height. Can be `"in"`, `"cm"` or `"px"`

- dpi:

  Image resolution used for conversion between units.

- address:

  An optional character string specifying an external URL, relative or
  absolute path to a file, or "mailto:" string (e.g.
  "mailto:example@example.com") that will be opened when the image is
  clicked.

- ...:

  additional arguments

## See also

[`wb_add_chart_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chart_xml.md)
[`wb_add_drawing()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_drawing.md)
[`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_mschart.md)
[`wb_add_plot()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_plot.md)

## Examples

``` r
img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")

wb <- wb_workbook()$
  add_worksheet()$
  add_image("Sheet 1", dims = "C5", file = img, width = 6, height = 5)$
  add_worksheet()$
  add_image(dims = "B2", file = img)$
  add_worksheet()$
  add_image(dims = "G3", file = img, width = 15, height = 12, units = "cm")
```
