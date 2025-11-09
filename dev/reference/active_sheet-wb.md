# Modify the state of active and selected sheets in a workbook

Get and set table of sheets and their state as selected and active in a
workbook

Multiple sheets can be selected, but only a single one can be active
(visible). The visible sheet, must not necessarily be a selected sheet.

## Usage

``` r
wb_get_active_sheet(wb)

wb_set_active_sheet(wb, sheet)

wb_get_selected(wb)

wb_set_selected(wb, sheet)
```

## Arguments

- wb:

  a workbook

- sheet:

  a sheet name of the workbook

## Value

a data frame with tabSelected and names

## Examples

``` r
wb <- wb_load(file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))
# testing is the selected sheet
wb_get_selected(wb)
#>   tabSelected topLeftCell workbookViewId  names
#> 1           1                          0 Sheet1
#> 2                      A3              0 Sheet2
# change the selected sheet to Sheet2
wb <- wb_set_selected(wb, "Sheet2")
# get the active sheet
wb_get_active_sheet(wb)
#> numeric(0)
# change the selected sheet to Sheet2
wb <- wb_set_active_sheet(wb, sheet = "Sheet2")
```
