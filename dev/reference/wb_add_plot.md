# Insert the current plot into a worksheet

The current plot is saved to a temporary image file using
[`grDevices::dev.copy()`](https://rdrr.io/r/grDevices/dev2.html) This
file is then written to the workbook using
[`wb_add_image()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_image.md).

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

  A workbook object

- sheet:

  A name or index of a worksheet

- dims:

  Worksheet dimension, single cell ("A1") or cell range ("A1:D4")

- width:

  Width of figure. Defaults to `6` in.

- height:

  Height of figure . Defaults to `4` in.

- row_offset, col_offset:

  Offset for column and row

- file_type:

  File type of image

- units:

  Units of width and height. Can be `"in"`, `"cm"` or `"px"`

- dpi:

  Image resolution

- ...:

  additional arguments

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
