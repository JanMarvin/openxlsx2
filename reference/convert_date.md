# Convert from spreadsheet date, datetime or hms number to R Date type

Convert from spreadsheet date number to R Date type

## Usage

``` r
convert_date(x, origin = "1900-01-01")

convert_datetime(x, origin = "1900-01-01", tz = "UTC")

convert_hms(x)
```

## Arguments

- x:

  A vector of integers

- origin:

  date. There are two options, 1900 or 1904. The default is what
  spreadsheet software usually uses

- tz:

  A timezone, defaults to "UTC"

## Value

A date, datetime, or hms.

## Details

Spreadsheet software stores dates as number of days from some origin day

Setting the timezone in `convert_datetime()` will alter the value. If
users expect a datetime value in a specific timezone, they should try
e.g. `lubridate::force_tz`.

## See also

[`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md)

## Examples

``` r
# date --
## 2014 April 21st to 25th
convert_date(c(41750, 41751, 41752, 41753, 41754, NA))
#> [1] "2014-04-21" "2014-04-22" "2014-04-23" "2014-04-24" "2014-04-25"
#> [6] NA          
convert_date(c(41750.2, 41751.99, NA, 41753))
#> [1] "2014-04-21" "2014-04-22" NA           "2014-04-24"

# datetime --
##  2014-07-01, 2014-06-30, 2014-06-29
x <- c(41821.8127314815, 41820.8127314815, NA, 41819, NaN)
convert_datetime(x)
#> [1] "2014-07-01 19:30:20 UTC" "2014-06-30 19:30:20 UTC"
#> [3] NA                        "2014-06-29 00:00:00 UTC"
#> [5] "NaN"                    
convert_datetime(x, tz = "Australia/Perth")
#> [1] "2014-07-02 03:30:20 AWST" "2014-07-01 03:30:20 AWST"
#> [3] NA                         "2014-06-29 08:00:00 AWST"
#> [5] "NaN"                     
convert_datetime(x, tz = "UTC")
#> [1] "2014-07-01 19:30:20 UTC" "2014-06-30 19:30:20 UTC"
#> [3] NA                        "2014-06-29 00:00:00 UTC"
#> [5] "NaN"                    

# hms ---
## 12:13:14
x <- 0.50918982
convert_hms(x)
#> [1] "12:13:14"
```
