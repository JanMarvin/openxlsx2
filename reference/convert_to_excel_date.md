# convert back to an Excel Date

convert back to an Excel Date

## Usage

``` r
convert_to_excel_date(df, date1904 = FALSE)
```

## Arguments

- df:

  dataframe

- date1904:

  take different origin

## Examples

``` r
 xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
 wb1 <- wb_load(xlsxFile)
 df <- wb_to_df(wb1)
 # conversion is done on dataframes only
 convert_to_excel_date(df = df["Var5"], date1904 = FALSE)
#>     Var5
#> 3  45075
#> 4  45069
#> 5  44958
#> 6     NA
#> 7     NA
#> 8  44987
#> 9     NA
#> 10 45284
#> 11 45285
#> 12 45138
```
