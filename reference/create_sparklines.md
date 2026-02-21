# Create a sparklines object

`create_sparklines()` defines a set of sparklines. Compact, word-sized
graphics that reside within a single cell. These are ideal for showing
trends in a series of values, such as seasonal increases or decreases,
or economic cycles, directly alongside the data.

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

  The name of the worksheet where the data originates.

- dims:

  A character string defining the source data range (e.g., "A1:E1").

- sqref:

  A character string defining the destination cell(s) (e.g., "F1").

- type:

  The type of sparkline: `NULL` (line), `"column"`, or `"stacked"`.

- negative:

  Logical; highlight negative data points.

- display_empty_cells_as:

  How to handle gaps in data: `"gap"`, `"span"` (connect points), or
  `"zero"`.

- markers:

  Logical; highlight all data points (Line type only).

- high, low, first, last:

  Logical; highlight the maximum, minimum, first, or last data points in
  the series.

- color_series, color_negative, color_axis, color_markers, color_first:

  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  objects defining the colors for various sparkline elements.

- color_last:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object for the color of the last point in the series.

- color_high:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object for the color of the highest point in the series.

- color_low:

  A
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/reference/wb_color.md)
  object for the color of the lowest point in the series.

- manual_max, manual_min:

  Numeric; optional fixed values for the y-axis.

- line_weight:

  Numeric; the thickness of the line (Line type only).

- date_axis:

  Logical; if `TRUE`, uses a date axis for the sparklines, allowing for
  irregular time intervals between data points.

- display_x_axis:

  Logical; show a horizontal axis.

- display_hidden:

  Logical; if `TRUE`, data in hidden rows or columns is plotted in the
  sparkline.

- min_axis_type, max_axis_type:

  Character; defines the scaling for the vertical axis. Options usually
  include "individual" (default), "group", or "custom".

- right_to_left:

  Logical; if `TRUE`, the sparkline is rendered from right to left.

- direction:

  The data orientation: `"row"` (default) or `"col"`. If `NULL`, the
  function attempts to infer direction from the dimensions.

- ...:

  Additional arguments.

## Value

A character string containing the XML structure for the sparkline group.

## Details

Sparklines are added to a workbook in "groups." A group shares the same
visual properties (type, colors, line weight, and axis settings). Within
a group, multiple individual sparklines are defined by pairing a data
range (`dims`) with a destination cell (`sqref`).

Types of Sparklines:

- `NULL` (Default): A standard line chart.

- `"column"`: A small column chart.

- `"stacked"`: Often referred to as a "Win/Loss" chart, where each data
  point is represented by a block indicating a positive or negative
  value.

Directionality: The `direction` argument determines how the `dims` range
is parsed. If you provide a multi-cell range like "A1:E10" as data for
10 sparklines, `direction = "row"` will treat each row as a separate
data series, while `direction = "col"` will treat each column as a
series.

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
