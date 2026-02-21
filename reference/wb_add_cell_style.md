# Modify the style in a cell region

The `wb_add_cell_style()` function provides direct access to the
cell-level formatting record (the `xf` node) within a `wbWorkbook`. It
is primarily used to control text alignment (horizontal and vertical),
text rotation, indentation, and cell protection (locking and hiding).

## Usage

``` r
wb_add_cell_style(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  apply_alignment = NULL,
  apply_border = NULL,
  apply_fill = NULL,
  apply_font = NULL,
  apply_number_format = NULL,
  apply_protection = NULL,
  border_id = NULL,
  ext_lst = NULL,
  fill_id = NULL,
  font_id = NULL,
  hidden = NULL,
  horizontal = NULL,
  indent = NULL,
  justify_last_line = NULL,
  locked = NULL,
  num_fmt_id = NULL,
  pivot_button = NULL,
  quote_prefix = NULL,
  reading_order = NULL,
  relative_indent = NULL,
  shrink_to_fit = NULL,
  text_rotation = NULL,
  vertical = NULL,
  wrap_text = NULL,
  xf_id = NULL,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet. Defaults to the current sheet.

- dims:

  A character string defining the cell range (e.g., "A1:K1").

- apply_alignment, apply_font, apply_fill, apply_border,
  apply_number_format, apply_protection:

  Logical; explicitly flags whether the spreadsheet software should
  apply the corresponding style category.

- ext_lst:

  Character; an optional XML string containing an extension list
  (`<extLst>`) for the cell style.

- font_id, fill_id, border_id, num_fmt_id:

  Optional; direct integer IDs referencing existing style sub-nodes.

- hidden:

  Logical; if `TRUE`, formulas are hidden when the sheet is protected.

- horizontal:

  Horizontal alignment. One of "general", "left", "center", "right",
  "fill", "justify", "centerContinuous", or "distributed".

- indent:

  Numeric; the indentation level for the cell content.

- justify_last_line:

  Logical; if `TRUE`, justifies the last line of text within the cell
  (useful for distributed alignment).

- locked:

  Logical; if `TRUE`, the cell cannot be edited when the sheet is
  protected.

- pivot_button:

  Logical; indicates if a pivot button should be displayed for the cell.

- quote_prefix:

  Logical; if `TRUE`, a single quote prefix is displayed in the formula
  bar but not the cell itself (often used for numbers stored as text).

- reading_order:

  Integer; the reading order for the cell content (e.g., 1 for
  Left-to-Right, 2 for Right-to-Left).

- relative_indent:

  Integer; the relative indentation level.

- shrink_to_fit:

  Logical; automatically reduces font size to fit the column width.

- text_rotation:

  Degrees of rotation (0 to 180).

- vertical:

  Vertical alignment. One of "top", "center", "bottom", "justify", or
  "distributed".

- wrap_text:

  Logical; enables line wrapping within the cell.

- xf_id:

  Integer; a direct reference to a master style (XF) ID in the style
  catalog.

- ...:

  Additional arguments.

## Value

The
[wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
object, invisibly.

The `wbWorkbook` object, invisibly

## Details

While functions like
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_font.md)
or
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_fill.md)
target specific sub-nodes of a style, `wb_add_cell_style()` manages the
properties that govern how content is positioned within the cell
boundaries and how it behaves when a worksheet is protected.

This function also allows for the direct assignment of style element IDs
(e.g., `font_id`, `fill_id`). This is an advanced feature that allows
users to map pre-existing styles in the workbook's style catalog to
specific cells.

Alignment and Text Control: Options such as `wrap_text`,
`shrink_to_fit`, and `text_rotation` are essential for managing
high-density data or creating stylized headers. The `text_rotation`
parameter accepts values in degrees (0–180), where values above 90
represent downward-slanting text.

Protection: The `locked` and `hidden` parameters only take effect when
worksheet protection is enabled (see
[`wb_protect_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_protect_worksheet.md)).
By default, all cells in a spreadsheet are "locked," but this has no
impact until the sheet is protected.

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_border.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_fill.md),
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_font.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_named_style.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_numfmt.md),
[`wb_cell_style`](https://janmarvin.github.io/openxlsx2/reference/wb_cell_style.md)

## Examples

``` r
wb <- wb_workbook()
wb <- wb_add_worksheet(wb, "S1")
wb <- wb_add_data(wb, "S1", x = mtcars)

wb <- wb_add_cell_style(
    wb,
    dims = "A1:K1",
    text_rotation = "45",
    horizontal = "center",
    vertical = "center",
    wrap_text = "1"
)
# Chaining
wb <- wb_workbook()$add_worksheet("S1")$add_data(x = mtcars)
wb$add_cell_style(dims = "A1:K1",
                  text_rotation = "45",
                  horizontal = "center",
                  vertical = "center",
                  wrap_text = "1")
```
