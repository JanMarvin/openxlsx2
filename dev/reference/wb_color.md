# Helper to create a color

Creates a `wbColour` object.

## Usage

``` r
wb_color(
  name = NULL,
  auto = NULL,
  indexed = NULL,
  hex = NULL,
  theme = NULL,
  tint = NULL,
  format = c("ARGB", "RGBA")
)
```

## Arguments

- name:

  A name of a color known to R either as name or RGB/ARGB/RGBA value.

- auto:

  A boolean.

- indexed:

  An indexed color value. This color has to be provided by the workbook.

- hex:

  A rgb color a RGB/ARGB/RGBA hex value with or without leading "#".

- theme:

  A zero based index referencing a value in the theme.

- tint:

  A tint value applied. Range from -1 (dark) to 1 (light).

- format:

  A colour format, one of ARGB (default) or RGBA.

## Value

a `wbColour` object

## Details

The **format** of the hex color representation can be either RGB, ARGB,
or RGBA. These hex formats differ only in a way how they encode the
transparency value alpha, ARGB expecting the alpha value before the RGB
values (default in spreadsheets), RGBA expects the alpha value after the
RGB values (default in R), and RGB is not encoding transparency at all.
If the colors some from functions such as `adjustcolor` that provide
color in the RGBA format, it is necessary to specify the
`format = "RGBA"` when calling the `wb_color()` function.

## See also

[`wb_get_base_colors()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_base_colors.md)
[`grDevices::colors()`](https://rdrr.io/r/grDevices/colors.html)
