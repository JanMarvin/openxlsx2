# Create border format

This function creates border styles for a cell in a spreadsheet. Border
styles can be any of the following: "none", "thin", "medium", "dashed",
"dotted", "thick", "double", "hair", "mediumDashed", "dashDot",
"mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot".
Border colors can be created with
[`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md).

## Usage

``` r
create_border(
  diagonal_down = "",
  diagonal_up = "",
  outline = "",
  bottom = NULL,
  bottom_color = NULL,
  diagonal = NULL,
  diagonal_color = NULL,
  end = "",
  horizontal = "",
  left = NULL,
  left_color = NULL,
  right = NULL,
  right_color = NULL,
  start = "",
  top = NULL,
  top_color = NULL,
  vertical = "",
  start_color = NULL,
  end_color = NULL,
  horizontal_color = NULL,
  vertical_color = NULL,
  ...
)
```

## Arguments

- diagonal_down, diagonal_up:

  Logical, whether the diagonal border goes from the bottom left to the
  top right, or top left to bottom right.

- outline:

  Logical, whether the border is.

- bottom, left, right, top, diagonal:

  Character, the style of the border.

- bottom_color, left_color, right_color, top_color, diagonal_color,
  start_color, end_color, horizontal_color, vertical_color:

  a
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md),
  the color of the border.

- horizontal, vertical:

  Character, the style of the inner border (only for dxf objects).

- start, end:

  leading and trailing edge of a border.

- ...:

  Additional arguments passed to other methods.

## Value

A formatted border object to be used in a spreadsheet.

## See also

[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_border.md)

Other style creating functions:
[`create_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_cell_style.md),
[`create_colors_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_colors_xml.md),
[`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_dxfs_style.md),
[`create_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_fill.md),
[`create_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_font.md),
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_numfmt.md),
[`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_tablestyle.md)

## Examples

``` r
# Create a border with a thick bottom and thin top
border <- create_border(
  bottom = "thick",
  bottom_color = wb_color("FF0000"),
  top = "thin",
  top_color = wb_color("00FF00")
)
```
