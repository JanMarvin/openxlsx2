# Remove all values in a worksheet

Remove content of a worksheet completely, or a region if specifying
`dims`.

## Usage

``` r
wb_clean_sheet(
  wb,
  sheet = current_sheet(),
  dims = NULL,
  numbers = TRUE,
  characters = TRUE,
  styles = TRUE,
  merged_cells = TRUE,
  hyperlinks = TRUE
)
```

## Arguments

- wb:

  A Workbook object

- sheet:

  sheet to clean

- dims:

  spreadsheet dimensions (optional)

- numbers:

  remove all numbers

- characters:

  remove all characters

- styles:

  remove all styles

- merged_cells:

  remove all merged_cells

- hyperlinks:

  remove all hyperlinks

## Value

A Workbook object
