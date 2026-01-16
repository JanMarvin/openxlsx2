# Modify font in a cell region

Modify the font in a cell region with more precision You can specify the
font in a cell with other cell styling functions, but `wb_add_font()`
gives you more control.

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

  A Workbook object

- sheet:

  the worksheet

- dims:

  the cell range

- name:

  Font name: default `"Aptos Narrow"`.

- color:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md),
  the color of the font. Default is "FF000000".

- size:

  Font size: default is `11`.

- bold:

  Logical, whether the font should be bold.

- italic:

  Logical, whether the font should be italic.

- outline:

  Logical, whether the font should have an outline.

- strike:

  Logical, whether the font should have a strikethrough.

- underline:

  underline, "single" or "double", default: ""

- charset:

  Character, the character set to be used. The list of valid IDs can be
  found in the **Details** section of
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md).

- condense:

  Logical, whether the font should be condensed.

- extend:

  Logical, whether the font should be extended.

- family:

  Character, the font family. Default is "2" (modern). "0" (auto), "1"
  (roman), "2" (swiss), "3" (modern), "4" (script), "5" (decorative). \#
  6-14 unused

- scheme:

  Character, the font scheme. Valid values are "minor", "major", "none".
  Default is "minor".

- shadow:

  Logical, whether the font should have a shadow.

- vert_align:

  Character, the vertical alignment of the font. Valid values are
  "baseline", "superscript", "subscript".

- update:

  Logical/Character if logical, all elements are assumed to be selected,
  whereas if character, only matching elements are updated. This will
  not alter strings styled with
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md).

- ...:

  ...

## Value

A `wbWorkbook`, invisibly

## Details

`wb_add_font()` provides all the options openxml accepts for a font
node, not all have to be set. Usually `name`, `size` and `color` should
be what the user wants. Setting `update` to `NULL` removes the style and
resets the cell to the workbook default.

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_border.md),
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_cell_style.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_fill.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_named_style.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_numfmt.md),
[`wb_cell_style`](https://janmarvin.github.io/openxlsx2/reference/wb_cell_style.md)

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
