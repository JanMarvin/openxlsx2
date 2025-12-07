# Get/set worksheet visible state in a workbook

Get and set worksheet visible state. This allows to hide worksheets from
the workbook. The visibility of a worksheet can either be "visible",
"hidden", or "veryHidden". You can set this when creating a worksheet
with `wb_add_worksheet(visible = FALSE)`

## Usage

``` r
wb_get_sheet_visibility(wb)

wb_set_sheet_visibility(wb, sheet = current_sheet(), value)
```

## Arguments

- wb:

  A `wbWorkbook` object

- sheet:

  Worksheet identifier

- value:

  a logical/character vector the same length as sheet, if providing a
  character vector, you can provide any of "hidden", "visible", or
  "veryHidden"

## Value

- `wb_set_sheet_visibility`: The Workbook object, invisibly.

- `wb_get_sheet_visibility()`: A character vector of the worksheet
  visibility value

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet(sheet = "S1", visible = FALSE)
wb$add_worksheet(sheet = "S2", visible = TRUE)
wb$add_worksheet(sheet = "S3", visible = FALSE)

wb$get_sheet_visibility()
#> [1] "hidden"  "visible" "hidden" 
wb$set_sheet_visibility(1, TRUE)         ## show sheet 1
wb$set_sheet_visibility(2, FALSE)        ## hide sheet 2
wb$set_sheet_visibility(3, "hidden")     ## hide sheet 3
wb$set_sheet_visibility(3, "veryHidden") ## hide sheet 3 from UI
```
