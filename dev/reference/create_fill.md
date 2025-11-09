# Create fill pattern

This function creates fill patterns for a cell in a spreadsheet. Fill
patterns can be simple solid colors or more complex gradient fills. For
certain pattern types, two colors are needed.

## Usage

``` r
create_fill(
  gradient_fill = "",
  pattern_type = "",
  bg_color = NULL,
  fg_color = NULL,
  ...
)
```

## Arguments

- gradient_fill:

  Character, specifying complex gradient fills.

- pattern_type:

  Character, specifying the fill pattern type. Valid values are "none"
  (default), "solid", "mediumGray", "darkGray", "lightGray",
  "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid",
  "darkTrellis", "lightHorizontal", "lightVertical", "lightDown",
  "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625".

- bg_color:

  Character, specifying the background color in hex8 format (alpha, red,
  green, blue) for pattern fills.

- fg_color:

  Character, specifying the foreground color in hex8 format (alpha, red,
  green, blue) for pattern fills.

- ...:

  Additional arguments passed to other methods.

## Value

A formatted fill pattern object to be used in a spreadsheet.

## See also

[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_fill.md)

Other style creating functions:
[`create_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_border.md),
[`create_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_cell_style.md),
[`create_colors_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_colors_xml.md),
[`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_dxfs_style.md),
[`create_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_font.md),
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_numfmt.md),
[`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_tablestyle.md)

## Examples

``` r
# Create a solid fill pattern with foreground color
fill <- create_fill(
  pattern_type = "solid",
  fg_color = wb_color(hex = "FFFF0000")
)
```
