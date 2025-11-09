# Set the default colors in a workbook

Modify / get the default colors of the workbook.

## Usage

``` r
wb_set_base_colors(wb, theme = "Office", ...)

wb_get_base_colors(wb, xml = FALSE, plot = TRUE)
```

## Arguments

- wb:

  A workbook object

- theme:

  a predefined color theme

- ...:

  optional parameters

- xml:

  Logical if xml string should be returned

- plot:

  Logical if a barplot of the colors should be returned

## Details

Theme must be any of the following: "Aspect", "Blue", "Blue II", "Blue
Green", "Blue Warm", "Greyscale", "Green", "Green Yellow", "Marquee",
"Median", "Office", "Office 2007 - 2010", "Office 2013 - 2022",
"Orange", "Orange Red", "Paper", "Red", "Red Orange", "Red Violet",
"Slipstream", "Violet", "Violet II", "Yellow", "Yellow Orange"

## See also

Other workbook styling functions:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`wb_add_dxfs_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_dxfs_style.md),
[`wb_add_style()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_style.md)

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chartsheet.md),
[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_worksheet.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_workbook.md)

## Examples

``` r
wb <- wb_workbook()
wb$get_base_colors()
#> $Office
#>      a:dk1      a:lt1      a:dk2      a:lt2  a:accent1  a:accent2  a:accent3 
#>  "#000000"  "#FFFFFF"  "#0E2841"  "#E8E8E8"  "#156082"  "#E97132"  "#196B24" 
#>  a:accent4  a:accent5  a:accent6    a:hlink a:folHlink 
#>  "#0F9ED5"  "#A02B93"  "#4EA72E"  "#467886"  "#96607D" 
#> 
wb$set_base_colors(theme = 3)
wb$set_base_colors(theme = "Violet II")
wb$get_base_colours()
#> $`Violet II`
#>      a:dk1      a:lt1      a:dk2      a:lt2  a:accent1  a:accent2  a:accent3 
#>  "#000000"  "#FFFFFF"  "#632E62"  "#EAE5EB"  "#92278F"  "#9B57D3"  "#755DD9" 
#>  a:accent4  a:accent5  a:accent6    a:hlink a:folHlink 
#>  "#665EB8"  "#45A5ED"  "#5982DB"  "#0066FF"  "#666699" 
#> 
```
