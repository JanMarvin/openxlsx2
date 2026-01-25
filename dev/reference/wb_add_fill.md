# Modify the background fill color in a cell region

The `wb_add_fill()` function applies background colors, patterns, or
gradients to a specified cell region. It allows for high-precision
styling, ranging from simple solid fills to complex geometric patterns
and linear or path-based gradients compliant with the OpenXML
specification.

## Usage

``` r
wb_add_fill(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  color = wb_color(hex = "FFFFFF00"),
  pattern = "solid",
  gradient_fill = "",
  every_nth_col = 1,
  every_nth_row = 1,
  bg_color = NULL,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet. Defaults to the current sheet.

- dims:

  A character string defining the cell range (e.g., "A1:D10").

- color:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md)
  object or hex string representing the primary fill (foreground) color.
  Defaults to yellow ("FFFFFF00").

- pattern:

  Character; the pattern type. Common values include "solid",
  "mediumGray", "lightGray", "darkGrid", and "lightTrellis". Defaults to
  "solid".

- gradient_fill:

  An optional XML string defining a gradient fill pattern. If provided,
  this overrides `color` and `pattern`.

- every_nth_col, every_nth_row:

  Numeric; applies the fill only to every \$n\$-th column or row within
  the specified `dims`. Useful for banding.

- bg_color:

  An optional
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md)
  for the background of a patterned fill.

- ...:

  Additional arguments.

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
object, invisibly.

The `wbWorkbook` object, invisibly

## Details

Background fills in spreadsheet software consist of a pattern type (the
most common being "solid") and a foreground color. If a non-solid
pattern is chosen (e.g., "darkVertical"), an optional `bg_color` can be
specified to create a two-tone effect.

The function also includes built-in logic for "nth" selection, which is
particularly useful for manual "zebra-striping" or creating grid-like
visual patterns without needing to manually construct a complex vector
of cell addresses.

Gradients: For advanced visual effects, `gradient_fill` accepts raw XML
strings defining `<gradientFill>` nodes. These can specify `degree` (for
linear gradients) or `type="path"` (for radial-style gradients) along
with multiple color stops.

Style Removal: Setting `color = NULL` removes the fill style from the
specified region, reverting the cells to the workbook's default
transparent background.

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_border.md),
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_cell_style.md),
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_font.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_named_style.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_numfmt.md),
[`wb_cell_style`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_cell_style.md)

## Examples

``` r
wb <- wb_workbook()
wb <- wb_add_worksheet(wb, "S1")
wb <- wb_add_data(wb, "S1", mtcars)
wb <- wb_add_fill(wb, "S1", dims = "D5:J23", color = wb_color(hex = "FFFFFF00"))
wb <- wb_add_fill(wb, "S1", dims = "B22:D27", color = wb_color(hex = "FF00FF00"))

wb <- wb_add_worksheet(wb, "S2")
wb <- wb_add_data(wb, "S2", mtcars)

gradient_fill1 <- '<gradientFill degree="90">
<stop position="0"><color rgb="FF92D050"/></stop>
<stop position="1"><color rgb="FF0070C0"/></stop>
</gradientFill>'
wb <- wb_add_fill(wb, "S2", dims = "A2:K5", gradient_fill = gradient_fill1)

gradient_fill2 <- '<gradientFill type="path" left="0.2" right="0.8" top="0.2" bottom="0.8">
<stop position="0"><color theme="0"/></stop>
<stop position="1"><color theme="4"/></stop>
</gradientFill>'
wb <- wb_add_fill(wb, "S2", dims = "A7:K10", gradient_fill = gradient_fill2)
```
