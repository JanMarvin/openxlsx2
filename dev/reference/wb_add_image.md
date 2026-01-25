# Insert an image into a worksheet

The `wb_add_image()` function embeds external image files into a
worksheet. It supports standard raster formats and provides granular
control over positioning through a variety of anchoring methods. Images
can be anchored to absolute positions, individual cells, or defined
ranges, and can optionally function as clickable hyperlinks.

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

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet to receive the image. Defaults to
  the current sheet.

- dims:

  A character string defining the placement. A single cell (e.g., "A1")
  uses a one-cell anchor; a range (e.g., "A1:D4") uses a two-cell
  anchor.

- file:

  The path to the image file. Supported formats are JPEG, PNG, and BMP.

- width, height:

  The numeric width and height of the image.

- row_offset, col_offset:

  Offset vectors for fine-tuning the position within the anchor cell(s).

- units:

  The units for `width` and `height`. Must be one of `"in"` (inches),
  `"cm"` (centimeters), or `"px"` (pixels).

- dpi:

  The resolution (dots per inch) used for conversion when `units` is set
  to `"px"`. Defaults to 300.

- address:

  An optional character string specifying a URL, file path, or "mailto:"
  link to be opened when the image is clicked.

- ...:

  Additional arguments. Includes support for legacy `start_row` and
  `start_col` parameters.

## Details

Image placement is determined by the `dims` argument and internal
anchoring logic. If a single cell is provided (e.g., "A1"), the image is
placed using a one-cell anchor where the top-left corner is fixed to the
cell. If a range is provided (e.g., "A1:D4"), a two-cell anchor is
utilized, which can cause the image to scale with the underlying rows
and columns.

Position offsets (`row_offset` and `col_offset`) allow for sub-cell
precision by shifting the image from its anchor point. Internally, all
dimensions are converted to English Metric Units (EMUs), where 1 inch
equals 914,400 EMUs, ensuring high-fidelity rendering across different
display scales.

Supported file types include `"jpeg"`, `"png"`, and `"bmp"`. If an
`address` is provided, the function creates a relationship to an
external target or email, transforming the image into a functional
hyperlink.

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
