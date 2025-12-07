# Remove a worksheet from a workbook

Remove a worksheet from a workbook

## Usage

``` r
wb_remove_worksheet(wb, sheet = current_sheet())
```

## Arguments

- wb:

  A wbWorkbook object

- sheet:

  The sheet name or index to remove

## Value

The `wbWorkbook` object, invisibly.

## Examples

``` r
## load a workbook
wb <- wb_load(file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))

## Remove sheet 2
wb <- wb_remove_worksheet(wb, 2)
```
