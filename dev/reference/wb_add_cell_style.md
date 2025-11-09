# Modify the style in a cell region

Add cell style to a cell region

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

  a workbook

- sheet:

  the worksheet

- dims:

  the cell range

- apply_alignment:

  logical apply alignment

- apply_border:

  logical apply border

- apply_fill:

  logical apply fill

- apply_font:

  logical apply font

- apply_number_format:

  logical apply number format

- apply_protection:

  logical apply protection

- border_id:

  border ID to apply

- ext_lst:

  extension list something like `<extLst>...</extLst>`

- fill_id:

  fill ID to apply

- font_id:

  font ID to apply

- hidden:

  logical cell is hidden

- horizontal:

  align content horizontal ('general', 'left', 'center', 'right',
  'fill', 'justify', 'centerContinuous', 'distributed')

- indent:

  logical indent content

- justify_last_line:

  logical justify last line

- locked:

  logical cell is locked

- num_fmt_id:

  number format ID to apply

- pivot_button:

  unknown

- quote_prefix:

  unknown

- reading_order:

  reading order left to right

- relative_indent:

  relative indentation

- shrink_to_fit:

  logical shrink to fit

- text_rotation:

  degrees of text rotation

- vertical:

  vertical alignment of content ('top', 'center', 'bottom', 'justify',
  'distributed')

- wrap_text:

  wrap text in cell

- xf_id:

  xf ID to apply

- ...:

  additional arguments

## Value

The `wbWorkbook` object, invisibly

## See also

Other styles:
[`wb_add_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_border.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_fill.md),
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_font.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_named_style.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_numfmt.md),
[`wb_cell_style`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_cell_style.md)

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
