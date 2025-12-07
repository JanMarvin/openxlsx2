# Create a custom formatting style

Create a new style to apply to worksheet cells. These styles are used in
conditional formatting and in (pivot) table styles.

## Usage

``` r
create_dxfs_style(
  font_name = NULL,
  font_size = NULL,
  font_color = NULL,
  num_fmt = NULL,
  border = NULL,
  border_color = wb_color(getOption("openxlsx2.borderColor", "black")),
  border_style = getOption("openxlsx2.borderStyle", "thin"),
  bg_fill = NULL,
  fg_color = NULL,
  gradient_fill = NULL,
  text_bold = NULL,
  text_strike = NULL,
  text_italic = NULL,
  text_underline = NULL,
  ...
)
```

## Arguments

- font_name:

  A name of a font. Note the font name is not validated. If `font_name`
  is `NULL`, the workbook `base_font` is used. (Defaults to Calibri),
  see
  [`wb_get_base_font()`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md)

- font_size:

  Font size. A numeric greater than 0. By default, the workbook base
  font size is used. (Defaults to 11)

- font_color:

  Color of text in cell. A valid hex color beginning with "#" or one of
  colors(). If `font_color` is NULL, the workbook base font colors is
  used. (Defaults to black)

- num_fmt:

  Cell formatting. Some custom openxml format

- border:

  `NULL` or `TRUE`

- border_color:

  "black"

- border_style:

  "thin"

- bg_fill:

  Cell background fill color.

- fg_color:

  Cell foreground fill color.

- gradient_fill:

  An xml string beginning with `<gradientFill>` ...

- text_bold:

  bold

- text_strike:

  strikeout

- text_italic:

  italic

- text_underline:

  underline 1, true, single or double

- ...:

  Additional arguments

## Value

A dxfs style node

## Details

It is possible to override border_color and border_style with {left,
right, top, bottom}\_color, {left, right, top, bottom}\_style.

## See also

[`wb_add_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_style.md)
[`wb_add_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_dxfs_style.md)

Other style creating functions:
[`create_border()`](https://janmarvin.github.io/openxlsx2/reference/create_border.md),
[`create_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/create_cell_style.md),
[`create_colors_xml()`](https://janmarvin.github.io/openxlsx2/reference/create_colors_xml.md),
[`create_fill()`](https://janmarvin.github.io/openxlsx2/reference/create_fill.md),
[`create_font()`](https://janmarvin.github.io/openxlsx2/reference/create_font.md),
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/create_numfmt.md),
[`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/reference/create_tablestyle.md)

## Examples

``` r
# do not apply anything
style1 <- create_dxfs_style()

# change font color and background color
style2 <- create_dxfs_style(
  font_color = wb_color(hex = "FF9C0006"),
  bg_fill = wb_color(hex = "FFFFC7CE")
)

# change font (type, size and color) and background
# the old default in openxlsx and openxlsx2 <= 0.3
style3 <- create_dxfs_style(
  font_name = "Aptos Narrow",
  font_size = 11,
  font_color = wb_color(hex = "FF9C0006"),
  bg_fill = wb_color(hex = "FFFFC7CE")
)

## See package vignettes for further examples
```
