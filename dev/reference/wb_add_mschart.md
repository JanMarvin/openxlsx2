# Add an mschart object to a worksheet

The `wb_add_mschart()` function allows for the seamless integration of
native charts created via the `mschart` package into a worksheet. Unlike
static images or plots, these are dynamic, native spreadsheet charts
that remain editable and can utilize data already present in the
workbook or data provided directly at creation.

## Usage

``` r
wb_add_mschart(
  wb,
  sheet = current_sheet(),
  dims = NULL,
  graph,
  col_offset = 0,
  row_offset = 0,
  ...
)
```

## Arguments

- wb:

  A
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object.

- sheet:

  The name or index of the worksheet where the chart will be placed.
  Defaults to the current sheet.

- dims:

  A character string defining the chart's position or range (e.g., "A1"
  or "F4:L20").

- graph:

  An `ms_chart` object created with the `mschart` package.

- col_offset, row_offset:

  Numeric values for fine-tuning the chart's displacement from its
  anchor point.

- ...:

  Additional arguments.

## Details

The function acts as a bridge between the `ms_chart` objects and the
spreadsheet's internal XML drawing structure. It interprets the chart
settings and data series to generate the necessary DrawingML.

There are two primary workflows for adding charts:

1.  External Data: If the `graph` object contains a standard data frame,
    `wb_add_mschart()` automatically writes this data to the worksheet
    before rendering the chart.

2.  Internal Data: If the `graph` object is initialized using a
    `wb_data` object (created via
    [`wb_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_data.md)),
    the chart will directly reference the existing cell ranges in the
    worksheet. This is the preferred method for maintaining a single
    source of truth for your data.

The chart is positioned using the `dims` argument. A single cell anchor
(e.g., "A1") will place the top-left corner of the chart, while a range
(e.g., "E5:L20") will scale the chart to fit that specific area.

## Notes

- This function requires the `mschart` package to be installed.

- Native charts are highly dependent on the calculation engine of the
  spreadsheet software; if the underlying data changes, the chart will
  update automatically when the file is opened.

- The function generates unique internal IDs for the chart axes to
  ensure compliance with the OpenXML specification.

## See also

[`wb_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_data.md)
[`wb_add_chart_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chart_xml.md)
[wb_add_image](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_image.md)
`wb_add_mschart()`
[wb_add_plot](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_plot.md)

## Examples

``` r
if (requireNamespace("mschart")) {
require(mschart)

## Add mschart to worksheet (adds data and chart)
scatter <- ms_scatterchart(data = iris, x = "Sepal.Length", y = "Sepal.Width", group = "Species")
scatter <- chart_settings(scatter, scatterstyle = "marker")

wb <- wb_workbook()
wb <- wb_add_worksheet(wb)
wb <- wb_add_mschart(wb, dims = "F4:L20", graph = scatter)

## Add mschart to worksheet and use available data
wb <- wb_workbook()
wb <- wb_add_worksheet(wb)
wb <- wb_add_data(wb, x = mtcars, dims = "B2")

# create wb_data object
dat <- wb_data(wb, 1, dims = "B2:E6")

# call ms_scatterplot
data_plot <- ms_scatterchart(
  data = dat,
  x = "mpg",
  y = c("disp", "hp"),
  labels = c("disp", "hp")
)

# add the scatterplot to the data
wb <- wb_add_mschart(wb, dims = "F4:L20", graph = data_plot)
}
#> Loading required namespace: mschart
#> Loading required package: mschart
```
