# Add a Chart object to a workbook sheet

Renders a `Chart` R6 object and injects the resulting XML into an
`openxlsx2` workbook at the specified location.

## Usage

``` r
wb_add_encharter(wb, sheet = current_sheet(), dims = NULL, graph)
```

## Arguments

- wb:

  An `openxlsx2` workbook object.

- sheet:

  Sheet name or index where the chart will be placed.

- dims:

  Character string defining the cell range (e.g., "E2:M20").

- graph:

  An initialized `Chart` R6 object.

## Value

The workbook object, invisibly.
