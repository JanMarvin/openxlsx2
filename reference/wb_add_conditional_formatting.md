# Add conditional formatting to cells in a worksheet

Add conditional formatting to cells. You can find more details in
[`vignette("conditional-formatting")`](https://janmarvin.github.io/openxlsx2/articles/conditional-formatting.md).

## Usage

``` r
wb_add_conditional_formatting(
  wb,
  sheet = current_sheet(),
  dims = NULL,
  rule = NULL,
  style = NULL,
  type = c("expression", "colorScale", "dataBar", "iconSet", "duplicatedValues",
    "uniqueValues", "containsErrors", "notContainsErrors", "containsBlanks",
    "notContainsBlanks", "containsText", "notContainsText", "beginsWith", "endsWith",
    "between", "topN", "bottomN"),
  params = list(showValue = TRUE, gradient = TRUE, border = TRUE, percent = FALSE, rank =
    5L, axisPosition = "automatic"),
  ...
)

wb_remove_conditional_formatting(
  wb,
  sheet = current_sheet(),
  dims = NULL,
  first = FALSE,
  last = FALSE
)
```

## Arguments

- wb:

  A Workbook object

- sheet:

  A name or index of a worksheet

- dims:

  A cell or cell range like "A1" or "A1:B2"

- rule:

  The condition under which to apply the formatting. See **Examples**.

- style:

  A name of a style to apply to those cells that satisfy the rule. See
  [`wb_add_dxfs_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_dxfs_style.md)
  how to create one. The default style has `font_color = "FF9C0006"` and
  `bg_fill = "FFFFC7CE"`

- type:

  The type of conditional formatting rule to apply. One of
  `"expression"`, `"colorScale"` or others mentioned in **Details**.

- params:

  A list of additional parameters passed. See **Details** for more.

- ...:

  additional arguments

- first:

  remove the first conditional formatting

- last:

  remove the last conditional formatting

## Details

openxml uses the alpha channel first then RGB, whereas the usual default
is RGBA.

Conditional formatting `type` accept different parameters. Unless noted,
unlisted parameters are ignored. If an expression is pointing to a cell
`"A1=1"`, this cell reference is fluid and not fixed like `"$A$1=1"`. It
will behave similar to a formula, when `dims` is spanning multiple
columns or rows (A1, A2, A3 ... in vertical direction, A1, B1, C1 ... in
horizontal direction). If `dims` is a non consecutive range
("A1:B2,D1:F2"), the expression is applied to each range. For the second
`dims` range it will be evaluated again as `"A1=1"`.

- `expression`:

  `[style]`  
  A `Style` object  
    
  `[rule]`  
  A formula expression (as a character). Valid operators are: `<`, `<=`,
  `>`, `>=`, `==`, `!=`

- `colorScale`:

  `[style]`  
  A `character` vector of valid colors with length `2` or `3`  
    
  `[rule]`  
  `NULL` or a `character` vector of valid colors of equal length to
  `styles`

- `dataBar`:

  `[style]`  
  A `character` vector of valid colors with length `2` or `3`  
    
  `[rule]`  
  A `numeric` vector specifying the range of the databar colors. Must be
  equal length to `style`  
    
  `[params$showValue]`  
  If `FALSE` the cell value is hidden. Default `TRUE`  
    
  `[params$gradient]`  
  If `FALSE` color gradient is removed. Default `TRUE`  
    
  `[params$border]`  
  If `FALSE` the border around the database is hidden. Default `TRUE`  
    
  `[params$direction]`  
  A `string` the direction in which the databar points. Must be equal to
  one of the following values: `"context"` (default), `"leftToRight"`,
  `"rightToLeft"`.  
    
  `[params${axisColor,borderColor,negativeBarColorSameAsPositive,negativeBarBorderColorSameAsPositive}]`
  Colors and bools configuring the style of the border.
  `[params$axisPosition]`  
  A `string` specifying the data bar's axis position. Must be equal to
  one of the following values: `"automatic"` (default, variable position
  based on negative values), `"middle"` (cell midpoint), `"none"`
  (negative bars shown in same direction as positive bars).  
    

- `duplicatedValues` / `uniqueValues` / `containsErrors`:

  `[style]`  
  A `Style` object

- `contains`:

  `[style]`  
  A `Style` object  
    
  `[rule]`  
  The text to look for within cells

- `between`:

  `[style]`  
  A `Style` object.  
    
  `[rule]`  
  A `numeric` vector of length `2` specifying lower and upper bound
  (Inclusive)

- `topN`:

  `[style]`  
  A `Style` object  
    
  `[params$rank]`  
  A `numeric` vector of length `1` indicating number of highest values.
  Default `5L`  
    
  `[params$percent]` If `TRUE`, uses percentage

- `bottomN`:

  `[style]`  
  A `Style` object  
    
  `[params$rank]`  
  A `numeric` vector of length `1` indicating number of lowest values.
  Default `5L`  
    
  `[params$percent]`  
  If `TRUE`, uses percentage

- `iconSet`:

  `[params$showValue]`  
  If `FALSE`, the cell value is hidden. Default `TRUE`  
    
  `[params$reverse]`  
  If `TRUE`, the order is reversed. Default `FALSE`  
    
  `[params$percent]`  
  If `TRUE`, uses percentage  
    
  `[params$iconSet]`  
  Uses one of the implemented icon sets. Values must match the length of
  the icons in the set 3Arrows, 3ArrowsGray, 3Flags, 3Signs, 3Stars,
  3Symbols, 3Symbols2, 3TrafficLights1, 3TrafficLights2, 3Triangles,
  4Arrows, 4ArrowsGray, 4Rating, 4RedToBlack, 4TrafficLights, 5Arrows,
  5ArrowsGray, 5Boxes, 5Quarters, 5Rating. The default is
  3TrafficLights1.

## See also

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/reference/col_widths-wb.md),
[`filter-wb`](https://janmarvin.github.io/openxlsx2/reference/filter-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/reference/grouping-wb.md),
[`named_region-wb`](https://janmarvin.github.io/openxlsx2/reference/named_region-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/reference/row_heights-wb.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_slicer.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_thread.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/reference/wb_merge_cells.md)

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("a")
wb$add_data(x = 1:4, col_names = FALSE)
wb$add_conditional_formatting(dims = wb_dims(cols = "A", rows = 1:4), rule = ">2")
```
