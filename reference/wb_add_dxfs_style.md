# Set a dxfs style for the workbook

The `wb_add_dxfs_style()` function defines a "Differential Formatting"
(DXF) style within a `wbWorkbook`. Unlike standard styles (XFs), which
are assigned directly to cells, DXF styles are used as templates for
dynamic formatting features such as conditional formatting rules and
custom table styles.

## Usage

``` r
wb_add_dxfs_style(
  wb,
  name,
  font_name = NULL,
  font_size = NULL,
  font_color = NULL,
  num_fmt = NULL,
  border = NULL,
  border_color = wb_color(getOption("openxlsx2.borderColor", "black")),
  border_style = getOption("openxlsx2.borderStyle", "thin"),
  bg_fill = NULL,
  gradient_fill = NULL,
  text_bold = NULL,
  text_italic = NULL,
  text_underline = NULL,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object.

- name:

  A unique character string to identify the DXF style.

- font_name:

  Character; the font name.

- font_size:

  Numeric; the font size.

- font_color:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object for the font.

- num_fmt:

  The number format string or ID.

- border:

  Logical; if `TRUE`, applies borders to the style.

- border_color:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object for the borders.

- border_style:

  Character; the border style (e.g., "thin", "thick"). Defaults to the
  "openxlsx2.borderStyle" option.

- bg_fill:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object for the background fill.

- gradient_fill:

  An optional XML string for a gradient fill pattern.

- text_bold:

  Logical; if `TRUE`, applies bold formatting.

- text_italic:

  Logical; if `TRUE`, applies italic formatting.

- text_underline:

  Logical; if `TRUE`, applies underline formatting.

- ...:

  Additional arguments passed to
  [`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md).

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
object, invisibly.

## Details

DXF styles are differential because they usually only define a subset of
cell properties (e.g., just the font color or a background fill). When a
conditional formatting rule is triggered, the properties defined in the
DXF style are layered on top of the cell's existing base style.

This function acts as a wrapper around
[`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md),
allowing you to bundle font, border, fill, and number format attributes
into a named style that can be referenced later by its `name`.

## See also

Other workbook styling functions:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/reference/base_font-wb.md),
[`wb_add_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_style.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/reference/wb_base_colors.md)

## Examples

``` r
wb <- wb_workbook()
wb <- wb_add_worksheet(wb)
wb <- wb_add_dxfs_style(
   wb,
   name = "nay",
   font_color = wb_color(hex = "FF9C0006"),
   bg_fill = wb_color(hex = "FFFFC7CE")
  )
```
