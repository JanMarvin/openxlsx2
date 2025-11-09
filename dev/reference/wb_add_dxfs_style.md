# Set a dxfs styling for the workbook

These styles are used with conditional formatting and custom table
styles.

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

  A Workbook object.

- name:

  the style name

- font_name:

  the font name

- font_size:

  the font size

- font_color:

  the font color (a
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md)
  object)

- num_fmt:

  the number format

- border:

  logical if borders are applied

- border_color:

  the border color

- border_style:

  the border style

- bg_fill:

  any background fill

- gradient_fill:

  any gradient fill

- text_bold:

  logical if text is bold

- text_italic:

  logical if text is italic

- text_underline:

  logical if text is underlined

- ...:

  additional arguments passed to
  [`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_dxfs_style.md)

## Value

The Workbook object, invisibly

## See also

Other workbook styling functions:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`wb_add_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_style.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_base_colors.md)

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
