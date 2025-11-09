# Modify the background fill color in a cell region

Add fill to a cell region.

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

  a workbook

- sheet:

  the worksheet

- dims:

  the cell range

- color:

  the colors to apply, e.g. yellow: wb_color(hex = "FFFFFF00")

- pattern:

  various default "none" but others are possible: "solid", "mediumGray",
  "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown",
  "darkUp", "darkGrid", "darkTrellis", "lightHorizontal",
  "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis",
  "gray125", "gray0625"

- gradient_fill:

  a gradient fill xml pattern.

- every_nth_col:

  which col should be filled

- every_nth_row:

  which row should be filled

- bg_color:

  (optional) background
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md)

- ...:

  ...

## Value

The `wbWorkbook` object, invisibly

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
