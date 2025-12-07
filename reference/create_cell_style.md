# Create cell style

This function creates a cell style for a spreadsheet, including
attributes such as borders, fills, fonts, and number formats.

## Usage

``` r
create_cell_style(
  border_id = "",
  fill_id = "",
  font_id = "",
  num_fmt_id = "",
  pivot_button = "",
  quote_prefix = "",
  xf_id = "",
  horizontal = "",
  indent = "",
  justify_last_line = "",
  reading_order = "",
  relative_indent = "",
  shrink_to_fit = "",
  text_rotation = "",
  vertical = "",
  wrap_text = "",
  ext_lst = "",
  hidden = "",
  locked = "",
  ...
)
```

## Arguments

- border_id, fill_id, font_id, num_fmt_id:

  IDs for style elements.

- pivot_button:

  Logical parameter for the pivot button.

- quote_prefix:

  Logical parameter for the quote prefix. (This way a number in a
  character cell will not cause a warning).

- xf_id:

  Dummy parameter for the xf ID. (Used only with named format styles).

- horizontal:

  Character, alignment can be ”, 'general', 'left', 'center', 'right',
  'fill', 'justify', 'centerContinuous', 'distributed'.

- indent:

  Integer parameter for the indent.

- justify_last_line:

  Logical for justifying the last line.

- reading_order:

  Logical parameter for reading order. 0 (Left to right; default) or 1
  (right to left).

- relative_indent:

  Dummy parameter for relative indent.

- shrink_to_fit:

  Logical parameter for shrink to fit.

- text_rotation:

  Integer parameter for text rotation (-180 to 180).

- vertical:

  Character, alignment can be ”, 'top', 'center', 'bottom', 'justify',
  'distributed'.

- wrap_text:

  Logical parameter for wrap text. (Required for linebreaks).

- ext_lst:

  Dummy parameter for extension list.

- hidden:

  Logical parameter for hidden.

- locked:

  Logical parameter for locked. (Impacts the cell only).

- ...:

  Reserved for additional arguments.

## Value

A formatted cell style object to be used in a spreadsheet.

## Details

A single cell style can make use of various other styles like border,
fill, and font. These styles are independent of the cell style and must
be registered with the style manager separately. This allows multiple
cell styles to share a common font type, for instance. The used style
elements are passed to the cell style via their IDs. An example of this
can be seen below. The number format can be a custom one created by
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/create_numfmt.md),
or a built-in style from the formats table below.

|      |                              |
|------|------------------------------|
| "ID" | "numFmt"                     |
| "0"  | "General"                    |
| "1"  | "0"                          |
| "2"  | "0.00"                       |
| "3"  | "#,##0"                      |
| "4"  | "#,##0.00"                   |
| "9"  | "0%"                         |
| "10" | "0.00%"                      |
| "11" | "0.00E+00"                   |
| "12" | "# ?/?"                      |
| "13" | "# ??/??"                    |
| "14" | "mm-dd-yy"                   |
| "15" | "d-mmm-yy"                   |
| "16" | "d-mmm"                      |
| "17" | "mmm-yy"                     |
| "18" | "h:mm AM/PM"                 |
| "19" | "h:mm:ss AM/PM"              |
| "20" | "h:mm"                       |
| "21" | "h:mm:ss"                    |
| "22" | "m/d/yy h:mm"                |
| "37" | "#,##0 ;(#,##0)"             |
| "38" | "#,##0 ;\[Red\](#,##0)"      |
| "39" | "#,##0.00;(#,##0.00)"        |
| "40" | "#,##0.00;\[Red\](#,##0.00)" |
| "45" | "mm:ss"                      |
| "46" | "\[h\]:mm:ss"                |
| "47" | "mmss.0"                     |
| "48" | "##0.0E+0"                   |
| "49" | "@"                          |

## See also

[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_cell_style.md)

Other style creating functions:
[`create_border()`](https://janmarvin.github.io/openxlsx2/reference/create_border.md),
[`create_colors_xml()`](https://janmarvin.github.io/openxlsx2/reference/create_colors_xml.md),
[`create_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/create_dxfs_style.md),
[`create_fill()`](https://janmarvin.github.io/openxlsx2/reference/create_fill.md),
[`create_font()`](https://janmarvin.github.io/openxlsx2/reference/create_font.md),
[`create_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/create_numfmt.md),
[`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/reference/create_tablestyle.md)

## Examples

``` r
foo_fill <- create_fill(pattern_type = "lightHorizontal",
                        fg_color = wb_color("blue"),
                        bg_color = wb_color("orange"))
foo_font <- create_font(sz = 36, b = TRUE, color = wb_color("yellow"))

wb <- wb_workbook()
wb$styles_mgr$add(foo_fill, "foo")
wb$styles_mgr$add(foo_font, "foo")

foo_style <- create_cell_style(
  fill_id = wb$styles_mgr$get_fill_id("foo"),
  font_id = wb$styles_mgr$get_font_id("foo")
)
```
