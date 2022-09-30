
<!-- README.md is generated from README.Rmd. Please edit that file -->

# openxlsx2

<!-- badges: start -->

[![R-CMD-check](https://github.com/JanMarvin/openxlsx2/workflows/R-CMD-check/badge.svg)](https://github.com/JanMarvin/openxlsx2/actions)
[![codecov](https://codecov.io/gh/JanMarvin/openxlsx2/branch/main/graph/badge.svg?token=HEZ7rXcZNq)](https://app.codecov.io/gh/JanMarvin/openxlsx2)
[![r-universe](https://janmarvin.r-universe.dev/badges/openxlsx2)](https://janmarvin.r-universe.dev/ui#package:openxlsx2)
<!-- badges: end -->

This R package is a modern reinterpretation of the widely used popular
`openxlsx` package. Similar to its predecessor, it simplifies the
creation of xlsx files by providing a clean interface for writing,
designing and editing worksheets. Based on a powerful XML library and
focusing on modern programming flows in pipes or chains, `openxlsx2`
allows to break many new ground.

Even though the project is already well progressed and supports most of
the features known and appreciated from the predecessor, there may still
be open gaps in one or the other place. A quick warning: Until the
stable version 1.0 there ~~may~~ will still be some changes to the API.

## Installation

You can install the stable version of `openxlsx2` with:

``` r
install.packages('openxlsx2')
```

You can install the development version of `openxlsx2` from
[GitHub](https://github.com/) with:

``` r
# install.packages("remotes")
remotes::install_github("JanMarvin/openxlsx2")
```

Or from [r-universe](https://r-universe.dev/) with:

``` r
# Enable repository from janmarvin
options(repos = c(
  janmarvin = 'https://janmarvin.r-universe.dev',
  CRAN = 'https://cloud.r-project.org'))
# Download and install openxlsx2 in R
install.packages('openxlsx2')
```

## Introduction

`openxlsx2` aims to be the swiss knife for working with the openxml
spreadsheet formats xlsx and xlsm (other formats of other spreadsheet
software are not supported). We offer two different variants how to work
with `openxlsx2`.

- The first one is to simply work with R objects. It is possible to read
  ([`read_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/read_xlsx.html))
  and write
  ([`write_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/write_xlsx.html))
  data from and to files. We offer a number of options in the commands
  to support various features of the openxml format, including reading
  and writing named ranges and tables. Furthermore, there are several
  ways to read certain information of an openxml spreadsheet without
  having opened it in a spreadsheet software before, e.g. to get the
  contained sheet names or tables.
- As a second variant `openxlsx2` offers the work with so called
  [`wbWorkbook`](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.html)
  objects. Here an openxml file is read into a corresponding
  `wbWorkbook` object
  ([`wb_load()`](https://janmarvin.github.io/openxlsx2/reference/wb_load.html))
  or a new one is created
  ([`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.html)).
  Afterwards the object can be further modified using various functions.
  For example, worksheets can be added or removed, the layout of cells
  or entire worksheets can be changed, and cells can be modified
  (overwritten or rewritten). Afterwards the `wbWorkbook` objects can be
  written as openxml files and processed by suitable spreadsheet
  software.

Many examples how to work with `openxlsx2` are in our [manual
pages](https://janmarvin.github.io/openxlsx2/reference/index.html) and
in our [vignettes](https://janmarvin.github.io/openxlsx2/articles/). You
can find them under:

``` r
vignette(package = "openxlsx2")
```

## Example

This is a basic example which shows you how to solve a common problem:

``` r
library(openxlsx2)
# read xlsx or xlsm files
path <- system.file("extdata/readTest.xlsx", package = "openxlsx2")
read_xlsx(path)
#>     Var1 Var2 NA  Var3  Var4       Var5         Var6    Var7
#> 2   TRUE    1 NA     1     a 2015-02-07 3209324 This #DIV/0!
#> 3   TRUE   NA NA #NUM!     b 2015-02-06         <NA>    <NA>
#> 4   TRUE    2 NA  1.34     c 2015-02-05         <NA>   #NUM!
#> 5  FALSE    2 NA  <NA> #NUM!       <NA>         <NA>    <NA>
#> 6  FALSE    3 NA  1.56     e       <NA>         <NA>    <NA>
#> 7  FALSE    1 NA   1.7     f 2015-02-02         <NA>    <NA>
#> 8     NA   NA NA  <NA>  <NA> 2015-02-01         <NA>    <NA>
#> 9  FALSE    2 NA    23     h 2015-01-31         <NA>    <NA>
#> 10 FALSE    3 NA  67.3     i 2015-01-30         <NA>    <NA>
#> 11    NA    1 NA   123  <NA> 2015-01-29         <NA>    <NA>

# or import workbooks
wb <- wb_load(path)
wb
#> A Workbook object.
#>  
#> Worksheets:
#>  Sheets: Sheet1 Sheet2 Sheet 3 Sheet 4 Sheet 5 Sheet 6 1 11 111 1111 11111 111111 
#>  Write order: 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12

# read a data frame
wb_to_df(wb)
#>     Var1 Var2 NA  Var3  Var4       Var5         Var6    Var7
#> 2   TRUE    1 NA     1     a 2015-02-07 3209324 This #DIV/0!
#> 3   TRUE   NA NA #NUM!     b 2015-02-06         <NA>    <NA>
#> 4   TRUE    2 NA  1.34     c 2015-02-05         <NA>   #NUM!
#> 5  FALSE    2 NA  <NA> #NUM!       <NA>         <NA>    <NA>
#> 6  FALSE    3 NA  1.56     e       <NA>         <NA>    <NA>
#> 7  FALSE    1 NA   1.7     f 2015-02-02         <NA>    <NA>
#> 8     NA   NA NA  <NA>  <NA> 2015-02-01         <NA>    <NA>
#> 9  FALSE    2 NA    23     h 2015-01-31         <NA>    <NA>
#> 10 FALSE    3 NA  67.3     i 2015-01-30         <NA>    <NA>
#> 11    NA    1 NA   123  <NA> 2015-01-29         <NA>    <NA>

# and save
temp <- temp_xlsx()
if (interactive()) wb_save(wb, temp)

## or create one yourself
wb <- wb_workbook()
# add a worksheet
wb$add_worksheet("sheet")
# add some data
wb$add_data("sheet", cars)
# open it in your default spreadsheet software
if (interactive()) wb$open()
```

## Authors and contributions

For a full list of all authors that have made this package possible and
for whom we are greatful, please see:

``` r
system.file("AUTHORS", package = "openxlsx2")
```

If you feel like you should be included on this list, please let us
know. If you have something to contribute, you are welcome. If something
is not working as expected, open issues or if you have solved an issue,
open a pull request. Please be respectful and be aware that we are
volunteers doing this for fun in our unpaid free time. We will work on
problems when we have time or need.

## License

This package is licensed under the MIT license and is based on
[`openxlsx`](https://github.com/ycphs/openxlsx) (by Alexander Walker and
Philipp Schauberger; COPYRIGHT 2014-2022) and
[`pugixml`](https://github.com/zeux/pugixml) (by Arseny Kapoulkine;
COPYRIGHT 2006-2022). Both released under the MIT license.
