# Add mschart object to a worksheet

Add mschart object to a worksheet

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

  a workbook

- sheet:

  the sheet on which the graph will appear

- dims:

  the dimensions where the sheet will appear

- graph:

  mschart object

- col_offset, row_offset:

  offsets for column and row

- ...:

  additional arguments

## See also

[`wb_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_data.md)
[`wb_add_chart_xml()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_chart_xml.md)
[wb_add_image](https://janmarvin.github.io/openxlsx2/reference/wb_add_image.md)
`wb_add_mschart()`
[wb_add_plot](https://janmarvin.github.io/openxlsx2/reference/wb_add_plot.md)

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
