# Create font format

This function creates font styles for a cell in a spreadsheet. It allows
customization of various font properties including bold, italic, color,
size, underline, and more.

## Usage

``` r
create_font(
  b = "",
  charset = "",
  color = wb_color(hex = "FF000000"),
  condense = "",
  extend = "",
  family = "2",
  i = "",
  name = "Aptos Narrow",
  outline = "",
  scheme = "minor",
  shadow = "",
  strike = "",
  sz = "11",
  u = "",
  vert_align = "",
  ...
)
```

## Arguments

- b:

  Logical, whether the font should be bold.

- charset:

  Character, the character set to be used. The list of valid IDs can be
  found in the **Details** section of
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/dev/reference/fmt_txt.md).

- color:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md),
  the color of the font. Default is "FF000000".

- condense:

  Logical, whether the font should be condensed.

- extend:

  Logical, whether the font should be extended.

- family:

  Character, the font family. Default is "2" (modern). "0" (auto), "1"
  (roman), "2" (swiss), "3" (modern), "4" (script), "5" (decorative). \#
  6-14 unused

- i:

  Logical, whether the font should be italic.

- name:

  Character, the name of the font. Default is "Aptos Narrow".

- outline:

  Logical, whether the font should have an outline.

- scheme:

  Character, the font scheme. Valid values are "minor", "major", "none".
  Default is "minor".

- shadow:

  Logical, whether the font should have a shadow.

- strike:

  Logical, whether the font should have a strikethrough.

- sz:

  Character, the size of the font. Default is "11".

- u:

  Character, the underline style. Valid values are "single", "double",
  "singleAccounting", "doubleAccounting", "none".

- vert_align:

  Character, the vertical alignment of the font. Valid values are
  "baseline", "superscript", "subscript".

- ...:

  Additional arguments passed to other methods.

## Value

A formatted font object to be used in a spreadsheet.

## See also

[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_font.md)

Other style creating functions:
[`create_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_border.md),
[`create_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_cell_style.md),
[`create_colors_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_colors_xml.md),
[`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_dxfs_style.md),
[`create_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_fill.md),
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_numfmt.md),
[`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_tablestyle.md)

## Examples

``` r
# Create a font with bold and italic styles
font <- create_font(
  b = TRUE,
  i = TRUE,
  color = wb_color(hex = "FF00FF00"),
  name = "Arial",
  sz = "12"
)

# openxml has the alpha value leading
hex8 <- unlist(xml_attr(read_xml(font), "font", "color"))
hex8 <- paste0("#", substr(hex8, 3, 8), substr(hex8, 1, 2))

# # write test color
# col <- crayon::make_style(col2rgb(hex8, alpha = TRUE))
# cat(col("Test"))
```
