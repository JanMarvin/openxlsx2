# Create sparklines object

Create a sparkline to be added a workbook with
[`wb_add_sparklines()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_sparklines.md)

## Usage

``` r
create_sparklines(
  sheet = current_sheet(),
  dims,
  sqref,
  type = NULL,
  negative = NULL,
  display_empty_cells_as = "gap",
  markers = NULL,
  high = NULL,
  low = NULL,
  first = NULL,
  last = NULL,
  color_series = wb_color(hex = "FF376092"),
  color_negative = wb_color(hex = "FFD00000"),
  color_axis = wb_color(hex = "FFD00000"),
  color_markers = wb_color(hex = "FFD00000"),
  color_first = wb_color(hex = "FFD00000"),
  color_last = wb_color(hex = "FFD00000"),
  color_high = wb_color(hex = "FFD00000"),
  color_low = wb_color(hex = "FFD00000"),
  manual_max = NULL,
  manual_min = NULL,
  line_weight = NULL,
  date_axis = NULL,
  display_x_axis = NULL,
  display_hidden = NULL,
  min_axis_type = NULL,
  max_axis_type = NULL,
  right_to_left = NULL,
  direction = NULL,
  ...
)
```

## Arguments

- sheet:

  sheet

- dims:

  Cell range of cells used to create the sparklines

- sqref:

  Cell range of the destination of the sparklines.

- type:

  Either `NULL`, `stacked` or `column`

- negative:

  negative

- display_empty_cells_as:

  Either `gap`, `span` or `zero`

- markers:

  markers add marker to line

- high:

  highlight highest value

- low:

  highlight lowest value

- first:

  highlight first value

- last:

  highlight last value

- color_series:

  colorSeries

- color_negative:

  colorNegative

- color_axis:

  colorAxis

- color_markers:

  colorMarkers

- color_first:

  colorFirst

- color_last:

  colorLast

- color_high:

  colorHigh

- color_low:

  colorLow

- manual_max:

  manualMax

- manual_min:

  manualMin

- line_weight:

  lineWeight

- date_axis:

  dateAxis

- display_x_axis:

  displayXAxis

- display_hidden:

  displayHidden

- min_axis_type:

  minAxisType

- max_axis_type:

  maxAxisType

- right_to_left:

  rightToLeft

- direction:

  Either `NULL`, `row` (or `1`) or `col` (or `2`). Should sparklines be
  created in the row or column direction? Defaults to `NULL`. When
  `NULL` the direction is inferred from `dims` in cases where `dims`
  spans a single row or column and defaults to `row` otherwise.

- ...:

  additional arguments

## Value

A string containing XML code

## Details

Colors are all predefined to be rgb. Maybe theme colors can be used too.

## Examples

``` r
# create multiple sparklines
sparklines <- c(
  create_sparklines("Sheet 1", "A3:L3", "M3", type = "column", first = "1"),
  create_sparklines("Sheet 1", "A2:L2", "M2", markers = "1"),
  create_sparklines("Sheet 1", "A4:L4", "M4", type = "stacked", negative = "1"),
  create_sparklines("Sheet 1", "A5:L5;A7:L7", "M5;M7", markers = "1")
)

t1 <- AirPassengers
t2 <- do.call(cbind, split(t1, cycle(t1)))
dimnames(t2) <- dimnames(.preformat.ts(t1))

wb <- wb_workbook()$
  add_worksheet("Sheet 1")$
  add_data(x = t2)$
  add_sparklines(sparklines = sparklines)

# create sparkline groups
sparklines <- c(
  create_sparklines("Sheet 2", "A2:L6;", "M2:M6", markers = "1"),
  create_sparklines(
    "Sheet 2", "A7:L7;A9:L9", "M7;M9", type = "stacked", negative = "1"
  ),
  create_sparklines(
    "Sheet 2", "A8:L8;A10:L13", "M8;M10:M13",
    type = "column", first = "1"
   ),
  create_sparklines(
    "Sheet 2", "A2:L13", "A14:L14", type = "column", first = "1",
    direction = "col"
  )
)

wb <- wb$
  add_worksheet("Sheet 2")$
  add_data(x = t2)$
  add_sparklines(sparklines = sparklines)
```
