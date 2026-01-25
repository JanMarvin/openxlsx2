# Modify borders in a cell region of a worksheet

The `wb_add_border()` function provides a high-level interface for
applying and managing cell borders within a `wbWorkbook`. It is designed
to handle both single cells and multi-cell regions, with built-in logic
to differentiate between exterior boundary borders and interior grid
lines.

## Usage

``` r
wb_add_border(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  bottom_color = wb_color(hex = "FF000000"),
  left_color = wb_color(hex = "FF000000"),
  right_color = wb_color(hex = "FF000000"),
  top_color = wb_color(hex = "FF000000"),
  bottom_border = "thin",
  left_border = "thin",
  right_border = "thin",
  top_border = "thin",
  inner_hgrid = NULL,
  inner_hcolor = NULL,
  inner_vgrid = NULL,
  inner_vcolor = NULL,
  update = FALSE,
  diagonal_up = NULL,
  diagonal_down = NULL,
  diagonal_color = NULL,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet to modify. Defaults to the current
  sheet.

- dims:

  A character string defining the cell range (e.g., "A1", "B2:G10").

- top_color, bottom_color, left_color, right_color:

  The colors for the exterior edges. Accepts
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md)
  objects or hex codes.

- top_border, bottom_border, left_border, right_border:

  The border style for the exterior edges of the range.

- inner_hgrid, inner_vgrid:

  The border style for internal horizontal and vertical grid lines
  within a range.

- inner_hcolor, inner_vcolor:

  The colors for internal grid lines.

- update:

  Logical or `NULL`. If `TRUE`, updates existing borders. If `NULL`,
  removes borders. If `FALSE` (default), overwrites existing styles with
  the new definition.

- diagonal_up, diagonal_down:

  Character string for the diagonal line style (e.g., "thin").

- diagonal_color:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md)
  object for the diagonal lines.

- ...:

  Additional arguments.

## Details

When applied to a range of cells (e.g., "A1:C3"), `wb_add_border()`
treats the selection as a single cohesive block. Parameters like
`top_border` and `left_border` apply only to the outermost edges of the
entire range. To draw lines between cells within the range, the
`inner_hgrid` (horizontal) and `inner_vgrid` (vertical) arguments are
used.

The function supports all standard spreadsheet border styles (e.g.,
"thin", "thick", "double", "dotted"). If `update = TRUE`, the function
attempts to merge new border definitions with existing ones, preserving
overlapping styles where possible. Setting `update = NULL` acts as a
reset, removing all border styles from the specified `dims` and
returning them to the workbook default.

For specialized needs, diagonal borders can be added using `diagonal_up`
and `diagonal_down`. Note that the OpenXML specification typically
restricts a cell to a single diagonal line style.

## Notes

- The function internally partitions the `dims` range into nine zones
  (corners, edges, and core) to apply the correct combination of
  exterior and interior borders efficiently.

- Color and style arguments must be paired; if a style is `NULL`, any
  assigned color for that side will be ignored.

- All border styles are registered in the workbook's global style
  catalog to ensure XML consistency.

## See also

[`create_border()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_border.md)

Other styles:
[`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_cell_style.md),
[`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_fill.md),
[`wb_add_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_font.md),
[`wb_add_named_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_named_style.md),
[`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_numfmt.md),
[`wb_cell_style`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_cell_style.md)

## Examples

``` r
wb <- wb_workbook()
wb <- wb_add_worksheet(wb, "S1")
wb <- wb_add_data(wb, "S1", mtcars)
wb <- wb_add_border(wb, 1, dims = "A1:K1",
 left_border = NULL, right_border = NULL,
 top_border = NULL, bottom_border = "double")
wb <- wb_add_border(wb, 1, dims = "A5",
 left_border = "dotted", right_border = "dotted",
 top_border = "hair", bottom_border = "thick")
wb <- wb_add_border(wb, 1, dims = "C2:C5")
wb <- wb_add_border(wb, 1, dims = "G2:H3")

wb <- wb_add_border(wb, 1, dims = "G12:H13",
 left_color = wb_color(hex = "FF9400D3"), right_color = wb_color(hex = "FF4B0082"),
 top_color = wb_color(hex = "FF0000FF"), bottom_color = wb_color(hex = "FF00FF00"))
wb <- wb_add_border(wb, 1, dims = "A20:C23")
wb <- wb_add_border(wb, 1, dims = "B12:D14",
 left_color = wb_color(hex = "FFFFFF00"), right_color = wb_color(hex = "FFFF7F00"),
 bottom_color = wb_color(hex = "FFFF0000"))
wb <- wb_add_border(wb, 1, dims = "D28:E28")

# With chaining

wb <- wb_workbook()
wb$add_worksheet("S1")$add_data("S1", mtcars)
wb$add_border(1, dims = "A1:K1",
 left_border = NULL, right_border = NULL,
 top_border = NULL, bottom_border = "double")
wb$add_border(1, dims = "A5",
 left_border = "dotted", right_border = "dotted",
 top_border = "hair", bottom_border = "thick")
wb$add_border(1, dims = "C2:C5")
wb$add_border(1, dims = "G2:H3")
wb$add_border(1, dims = "G12:H13",
 left_color = wb_color(hex = "FF9400D3"), right_color = wb_color(hex = "FF4B0082"),
 top_color = wb_color(hex = "FF0000FF"), bottom_color = wb_color(hex = "FF00FF00"))
wb$add_border(1, dims = "A20:C23")
wb$add_border(1, dims = "B12:D14",
 left_color = wb_color(hex = "FFFFFF00"), right_color = wb_color(hex = "FFFF7F00"),
 bottom_color = wb_color(hex = "FFFF0000"))
wb$add_border(1, dims = "D28:E28")
# if (interactive()) wb$open()

wb <- wb_workbook()
wb$add_worksheet("S1")$add_data("S1", mtcars)
wb$add_border(1, dims = "A2:K33", inner_vgrid = "thin",
 inner_vcolor = wb_color(hex = "FF808080"))

wb$add_worksheet()$
  add_border(dims = "B2:D4", bottom_border = "thick", left_border = "thick",
    right_border = "thick", top_border = "thick")$
  add_border(dims = "C3:E5", update = TRUE)

wb$add_worksheet()$
  add_border(
    dims = "B2:D4",
    diagonal_up = "thin",
    diagonal_down = "thin",
    diagonal_color = wb_color("red")
  )
```
