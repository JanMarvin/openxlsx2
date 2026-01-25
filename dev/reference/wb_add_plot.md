# Insert the current R plot into a worksheet

The
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_worksheet.md)
function captures the active R graphics device and embeds the displayed
plot into a worksheet. This is achieved by copying the current plot to a
temporary image file via
[`grDevices::dev.copy()`](https://rdrr.io/r/grDevices/dev2.html) and
subsequently invoking
[`wb_add_image()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_image.md)
to handle the workbook integration.

## Usage

``` r
wb_add_plot(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  width = 6,
  height = 4,
  row_offset = 0,
  col_offset = 0,
  file_type = "png",
  units = "in",
  dpi = 300,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet where the plot will be inserted.
  Defaults to the current sheet.

- dims:

  A character string defining the anchor point or range (e.g., "A1" or
  "A1:D4").

- width, height:

  The numeric dimensions of the exported plot. Defaults to 6x4 inches.

- row_offset, col_offset:

  Numeric vectors for sub-cell positioning offsets.

- file_type:

  The image format for the temporary capture. Supported types include
  `"png"`, `"jpeg"`, `"tiff"`, and `"bmp"`.

- units:

  The measurement units for `width` and `height`. Must be one of `"in"`,
  `"cm"`, or `"px"`.

- dpi:

  The resolution in dots per inch for the image conversion.

- ...:

  Additional arguments. Supports the deprecated `start_row` and
  `start_col` parameters for backward compatibility.

## Details

Because this function relies on the active graphics device, a plot must
be currently displayed in the R session (e.g., in the Plots pane or a
separate window) for the capture to succeed. The function supports
various file formats for the intermediate transition, including `"png"`,
`"jpeg"`, `"tiff"`, and `"bmp"`.

Positioning is managed through the spreadsheet coordinate system. Using
a single cell in `dims` (e.g., "A1") establishes a one-cell anchor where
the plot maintains its absolute dimensions. Providing a range (e.g.,
"A1:E10") creates a two-cell anchor, which may result in the plot
resizing if columns or rows within that range are adjusted in
spreadsheet software.

For programmatic control over the output quality, the `dpi` argument
influences the resolution of the captured device. Users working with
high-resolution displays or requiring print-quality outputs should
adjust the `dpi` and `units` accordingly.

## See also

[`wb_add_chart_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chart_xml.md)
[`wb_add_drawing()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_drawing.md)
[`wb_add_image()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_image.md)
[`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_mschart.md)

## Examples

``` r
if (requireNamespace("ggplot2") && interactive()) {
## Create a new workbook
wb <- wb_workbook()

## Add a worksheet
wb$add_worksheet("Sheet 1", grid_lines = FALSE)

## create plot objects
require(ggplot2)
p1 <- ggplot(mtcars, aes(x = mpg, fill = as.factor(gear))) +
  ggtitle("Distribution of Gas Mileage") +
  geom_density(alpha = 0.5)
p2 <- ggplot(Orange, aes(x = age, y = circumference, color = Tree)) +
  geom_point() + geom_line()

## Insert currently displayed plot to sheet 1, row 1, column 1
print(p1) # plot needs to be showing
wb$add_plot(1, width = 5, height = 3.5, file_type = "png", units = "in")

## Insert plot 2
print(p2)
wb$add_plot(1, dims = "J2", width = 16, height = 10, file_type = "png", units = "cm")

}
#> Loading required namespace: ggplot2
```
