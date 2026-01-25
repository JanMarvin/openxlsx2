# Modify font properties in a cell region

The `wb_add_font()` function provides granular control over the visual
appearance of text within a specified cell region. While other styling
functions include basic font options, `wb_add_font()` exposes the full
range of font attributes supported by the OpenXML specification,
allowing for precise adjustments to typeface, sizing, color, and
emphasis.

## Usage

``` r
wb_add_font(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  name = "Aptos Narrow",
  color = wb_color(hex = "FF000000"),
  size = "11",
  bold = "",
  italic = "",
  outline = "",
  strike = "",
  underline = "",
  charset = "",
  condense = "",
  extend = "",
  family = "",
  scheme = "",
  shadow = "",
  vert_align = "",
  update = FALSE,
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

  A character string defining the cell range (e.g., "A1:K1").

- name:

  Character; the font name. Defaults to "Aptos Narrow".

- color:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md)
  object or hex string defining the font color. Defaults to black
  ("FF000000").

- size:

  Numeric; the font size. Defaults to 11.

- bold:

  Logical; applies bold formatting if `TRUE`.

- italic:

  Logical; applies italic formatting if `TRUE`.

- outline:

  Logical; applies an outline effect to the text.

- strike:

  Logical; applies a strikethrough effect.

- underline:

  Character; the underline style, such as "single" or "double".

- charset:

  Character; the character set ID. See
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/dev/reference/fmt_txt.md)
  for details.

- condense:

  Logical; whether the font should be condensed.

- extend:

  Logical; whether the font should be extended.

- family:

  Character; the font family index (e.g., "1" for Roman, "2" for Swiss).

- scheme:

  Character; the font scheme. One of "minor", "major", or "none".

- shadow:

  Logical; applies a shadow effect to the text.

- vert_align:

  Character; vertical alignment. Options are "baseline", "superscript",
  or "subscript".

- update:

  Logical or character vector. Controls whether to overwrite the entire
  font style or only update specific properties.

- ...:

  Additional arguments.

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
object, invisibly.

A `wbWorkbook`, invisibly

## Details

This function operates on the font node of a cell's style. It is
particularly powerful when used with the `update` argument, which allows
users to modify specific attributes (like color) while preserving other
existing font properties (like bold or font name).

For common tasks, adjusting `name`, `size`, and `color` is sufficient.
However, the function also supports advanced properties like
`vert_align` (for subscripts/superscripts), `family` (font categories),
and `scheme` (theme-based font sets).

Note on Updates:

- If `update = FALSE` (default), the function applies the new font
  definition as a complete replacement for the existing font style.

- If `update` is a character vector (e.g., `c("color", "size")`), only
  those specific attributes are modified, and all other existing font
  properties are retained.

- Setting `update = NULL` removes the custom font style entirely,
  reverting the cells to the workbook's default font.

## Notes

- This function modifies the cell-level style and does not alter rich
  text strings created with
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/dev/reference/fmt_txt.md).

- Font styles are pooled in the workbook's style manager to ensure
  efficiency and XML compliance.

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_border.md),
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_cell_style.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_fill.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_named_style.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_numfmt.md),
[`wb_cell_style`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_cell_style.md)

## Examples

``` r
 wb <- wb_workbook()
 wb <- wb_add_worksheet(wb, "S1")
 wb <- wb_add_data(wb, "S1", mtcars)
 wb <- wb_add_font(wb, "S1", "A1:K1", name = "Arial", color = wb_color(theme = "4"))
# With chaining
 wb <- wb_workbook()$add_worksheet("S1")$add_data("S1", mtcars)
 wb$add_font("S1", "A1:K1", name = "Arial", color = wb_color(theme = "4"))

# Update the font color
 wb$add_font("S1", "A1:K1", color = wb_color("orange"), update = c("color"))
```
