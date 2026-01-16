# Create a data frame from a Workbook

Simple function to create a `data.frame` from a sheet in workbook.
Simple as in it was simply written down. `read_xlsx()` and `wb_read()`
are just internal wrappers of `wb_to_df()` intended for people coming
from other packages.

## Usage

``` r
wb_to_df(
  file,
  sheet,
  start_row = NULL,
  start_col = NULL,
  row_names = FALSE,
  col_names = TRUE,
  skip_empty_rows = FALSE,
  skip_empty_cols = FALSE,
  skip_hidden_rows = FALSE,
  skip_hidden_cols = FALSE,
  rows = NULL,
  cols = NULL,
  detect_dates = TRUE,
  na = "#N/A",
  fill_merged_cells = FALSE,
  dims,
  show_formula = FALSE,
  convert = TRUE,
  types,
  named_region,
  keep_attributes = FALSE,
  check_names = FALSE,
  show_hyperlinks = FALSE,
  apply_numfmts = FALSE,
  ...
)

read_xlsx(
  file,
  sheet,
  start_row = NULL,
  start_col = NULL,
  row_names = FALSE,
  col_names = TRUE,
  skip_empty_rows = FALSE,
  skip_empty_cols = FALSE,
  rows = NULL,
  cols = NULL,
  detect_dates = TRUE,
  named_region,
  na = "#N/A",
  fill_merged_cells = FALSE,
  check_names = FALSE,
  show_hyperlinks = FALSE,
  ...
)

wb_read(
  file,
  sheet = 1,
  start_row = NULL,
  start_col = NULL,
  row_names = FALSE,
  col_names = TRUE,
  skip_empty_rows = FALSE,
  skip_empty_cols = FALSE,
  rows = NULL,
  cols = NULL,
  detect_dates = TRUE,
  named_region,
  na = "NA",
  check_names = FALSE,
  show_hyperlinks = FALSE,
  ...
)
```

## Arguments

- file:

  An xlsx file,
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md)
  object or URL to xlsx file.

- sheet:

  Either sheet name or index. When missing the first sheet in the
  workbook is selected.

- start_row:

  first row to begin looking for data.

- start_col:

  first column to begin looking for data.

- row_names:

  If `TRUE`, the first col of data will be used as row names.

- col_names:

  If `TRUE`, the first row of data will be used as column names.

- skip_empty_rows:

  If `TRUE`, empty rows are skipped.

- skip_empty_cols:

  If `TRUE`, empty columns are skipped.

- skip_hidden_rows:

  If `TRUE`, hidden rows are skipped.

- skip_hidden_cols:

  If `TRUE`, hidden columns are skipped.

- rows:

  A numeric vector specifying which rows in the xlsx file to read. If
  `NULL`, all rows are read.

- cols:

  A numeric vector specifying which columns in the xlsx file to read. If
  `NULL`, all columns are read.

- detect_dates:

  If `TRUE`, attempt to recognize dates and perform conversion.

- na:

  Defines values to be treated as NA. Can be a character vector of
  strings or a named list: list(strings = ..., numbers = ...). Blank
  cells are always converted to `NA`.

- fill_merged_cells:

  If `TRUE`, the value in a merged cell is given to all cells within the
  merge.

- dims:

  Character string of type "A1:B2" as optional dimensions to be
  imported.

- show_formula:

  If `TRUE`, the underlying spreadsheet formulas are shown.

- convert:

  If `TRUE`, a conversion to dates and numerics is attempted.

- types:

  A named numeric indicating, the type of the data. Names must match the
  returned data. See **Details** for more.

- named_region:

  Character string with a `named_region` (defined name or table). If no
  sheet is selected, the first appearance will be selected. See
  [`wb_get_named_regions()`](https://janmarvin.github.io/openxlsx2/reference/named_region-wb.md)

- keep_attributes:

  If `TRUE` additional attributes are returned. (These are used
  internally to define a cell type.)

- check_names:

  If `TRUE` then the names of the variables in the data frame are
  checked to ensure that they are syntactically valid variable names.

- show_hyperlinks:

  If `TRUE` instead of the displayed text, hyperlink targets are shown.

- apply_numfmts:

  If `TRUE` numeric formats are applied if detected.

- ...:

  additional arguments

## Details

The returned data frame will have named rows matching the rows of the
worksheet. With `col_names = FALSE` the returned data frame will have
column names matching the columns of the worksheet. Otherwise the first
row is selected as column name.

Depending if the R package `hms` is loaded, `wb_to_df()` returns `hms`
variables or string variables in the `hh:mm:ss` format.

The `types` argument can be a named numeric or a character string of the
matching R variable type. Either `c(foo = 1)` or `c(foo = "numeric")`.

- 0: character

- 1: numeric

- 2: Date

- 3: POSIXct (datetime)

- 4: logical

If no type is specified, the column types are derived based on all cells
in a column within the selected data range, excluding potential column
names. If `keep_attr` is `TRUE`, the derived column types can be
inspected as an attribute of the data frame.

`wb_to_df()` will not pick up formulas added to a workbook object via
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_formula.md).
This is because only the formula is written and left to be evaluated
when the file is opened in a spreadsheet software. Opening, saving and
closing the file in a spreadsheet software will resolve this.

Before release 1.15, datetime variables (in 'yyyy-mm-dd hh:mm:ss'
format) were imported using the user's local system timezone
([`Sys.timezone()`](https://rdrr.io/r/base/timezones.html)). This
behavior has been updated. Now, all datetime variables are imported with
the timezone set to "UTC". If automatic date detection and conversion
are enabled but the conversion is unsuccessful (for instance, in a
column containing a mix of data types like strings, numbers, and dates)
dates might be displayed as a Unix timestamp. Usually they are converted
to character for character columns. If date detection is disabled, dates
will show up as a spreadsheet date format. To convert these, you can use
the functions
[`convert_date()`](https://janmarvin.github.io/openxlsx2/reference/convert_date.md),
[`convert_datetime()`](https://janmarvin.github.io/openxlsx2/reference/convert_date.md),
or
[`convert_hms()`](https://janmarvin.github.io/openxlsx2/reference/convert_date.md).
If types are specified, date detection is disabled.

You can use wildcards for all available columns or rows in `dims` by
using `+` and `-`. For example, `dims = "A-:+9"` will read everything
from the first row in column A through the last column in row 9. This
makes it unnecessary to update dimensions when working with files whose
sizes change frequently.

The function to apply numeric formats was not extensively tested for
numeric equality with spreadsheet software. There might be differences
and the function has limited support for builtin styles.

## See also

[`wb_get_named_regions()`](https://janmarvin.github.io/openxlsx2/reference/named_region-wb.md),
[openxlsx2](https://janmarvin.github.io/openxlsx2/reference/openxlsx2-package.md)

## Examples

``` r
###########################################################################
# numerics, dates, missings, bool and string
example_file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
wb1 <- wb_load(example_file)

# import workbook
wb_to_df(wb1)
#>     Var1 Var2 <NA>  Var3  Var4       Var5         Var6    Var7     Var8
#> 3   TRUE    1   NA     1     a 2023-05-29 3209324 This #DIV/0! 01:27:15
#> 4   TRUE   NA   NA #NUM!     b 2023-05-23         <NA>       0 14:02:57
#> 5   TRUE    2   NA  1.34     c 2023-02-01         <NA> #VALUE! 23:01:02
#> 6  FALSE    2   NA  <NA> #NUM!       <NA>         <NA>       2 17:24:53
#> 7  FALSE    3   NA  1.56     e       <NA>         <NA>    <NA>     <NA>
#> 8  FALSE    1   NA   1.7     f 2023-03-02         <NA>     2.7 08:45:58
#> 9     NA   NA   NA  <NA>  <NA>       <NA>         <NA>    <NA>     <NA>
#> 10 FALSE    2   NA    23     h 2023-12-24         <NA>      25     <NA>
#> 11 FALSE    3   NA  67.3     i 2023-12-25         <NA>       3     <NA>
#> 12    NA    1   NA   123  <NA> 2023-07-31         <NA>     122     <NA>

# do not convert first row to column names
wb_to_df(wb1, col_names = FALSE)
#>        B    C  D     E     F          G            H       I        J
#> 2   Var1 Var2 NA  Var3  Var4       Var5         Var6    Var7     Var8
#> 3   TRUE    1 NA     1     a 2023-05-29 3209324 This #DIV/0! 01:27:15
#> 4   TRUE <NA> NA #NUM!     b 2023-05-23         <NA>       0 14:02:57
#> 5   TRUE    2 NA  1.34     c 2023-02-01         <NA> #VALUE! 23:01:02
#> 6  FALSE    2 NA  <NA> #NUM!       <NA>         <NA>       2 17:24:53
#> 7  FALSE    3 NA  1.56     e       <NA>         <NA>    <NA>     <NA>
#> 8  FALSE    1 NA   1.7     f 2023-03-02         <NA>     2.7 08:45:58
#> 9   <NA> <NA> NA  <NA>  <NA>       <NA>         <NA>    <NA>     <NA>
#> 10 FALSE    2 NA    23     h 2023-12-24         <NA>      25     <NA>
#> 11 FALSE    3 NA  67.3     i 2023-12-25         <NA>       3     <NA>
#> 12  <NA>    1 NA   123  <NA> 2023-07-31         <NA>     122     <NA>

# do not try to identify dates in the data
wb_to_df(wb1, detect_dates = FALSE)
#>     Var1 Var2 <NA>  Var3  Var4  Var5         Var6    Var7       Var8
#> 3   TRUE    1   NA     1     a 45075 3209324 This #DIV/0! 0.06059028
#> 4   TRUE   NA   NA #NUM!     b 45069         <NA>       0 0.58538194
#> 5   TRUE    2   NA  1.34     c 44958         <NA> #VALUE! 0.95905093
#> 6  FALSE    2   NA  <NA> #NUM!    NA         <NA>       2 0.72561343
#> 7  FALSE    3   NA  1.56     e    NA         <NA>    <NA>         NA
#> 8  FALSE    1   NA   1.7     f 44987         <NA>     2.7 0.36525463
#> 9     NA   NA   NA  <NA>  <NA>    NA         <NA>    <NA>         NA
#> 10 FALSE    2   NA    23     h 45284         <NA>      25         NA
#> 11 FALSE    3   NA  67.3     i 45285         <NA>       3         NA
#> 12    NA    1   NA   123  <NA> 45138         <NA>     122         NA

# return the underlying spreadsheet formula instead of their values
wb_to_df(wb1, show_formula = TRUE)
#>     Var1 Var2 <NA>  Var3  Var4       Var5         Var6            Var7     Var8
#> 3   TRUE    1   NA     1     a 2023-05-29 3209324 This            E3/0 01:27:15
#> 4   TRUE   NA   NA #NUM!     b 2023-05-23         <NA>              C4 14:02:57
#> 5   TRUE    2   NA  1.34     c 2023-02-01         <NA>         #VALUE! 23:01:02
#> 6  FALSE    2   NA  <NA> #NUM!       <NA>         <NA>           C6+E6 17:24:53
#> 7  FALSE    3   NA  1.56     e       <NA>         <NA>            <NA>     <NA>
#> 8  FALSE    1   NA   1.7     f 2023-03-02         <NA>           C8+E8 08:45:58
#> 9     NA   NA   NA  <NA>  <NA>       <NA>         <NA>            <NA>     <NA>
#> 10 FALSE    2   NA    23     h 2023-12-24         <NA>    SUM(C10,E10)     <NA>
#> 11 FALSE    3   NA  67.3     i 2023-12-25         <NA> PRODUCT(C11,E3)     <NA>
#> 12    NA    1   NA   123  <NA> 2023-07-31         <NA>         E12-C12     <NA>

# read dimension without colNames
wb_to_df(wb1, dims = "A2:C5", col_names = FALSE)
#>    A    B    C
#> 2 NA Var1 Var2
#> 3 NA TRUE    1
#> 4 NA TRUE <NA>
#> 5 NA TRUE    2

# read selected cols
wb_to_df(wb1, cols = c("A:B", "G"))
#>    <NA>  Var1       Var5
#> 3    NA  TRUE 2023-05-29
#> 4    NA  TRUE 2023-05-23
#> 5    NA  TRUE 2023-02-01
#> 6    NA FALSE       <NA>
#> 7    NA FALSE       <NA>
#> 8    NA FALSE 2023-03-02
#> 9    NA    NA       <NA>
#> 10   NA FALSE 2023-12-24
#> 11   NA FALSE 2023-12-25
#> 12   NA    NA 2023-07-31

# read selected rows
wb_to_df(wb1, rows = c(2, 4, 6))
#>    Var1 Var2 <NA>  Var3  Var4       Var5 Var6 Var7     Var8
#> 4  TRUE   NA   NA #NUM!     b 2023-05-23   NA    0 14:02:57
#> 6 FALSE    2   NA  <NA> #NUM!       <NA>   NA    2 17:24:53

# convert characters to numerics and date (logical too?)
wb_to_df(wb1, convert = FALSE)
#>     Var1 Var2 <NA>  Var3  Var4       Var5         Var6    Var7     Var8
#> 3   TRUE    1 <NA>     1     a 2023-05-29 3209324 This #DIV/0! 01:27:15
#> 4   TRUE <NA> <NA> #NUM!     b 2023-05-23         <NA>       0 14:02:57
#> 5   TRUE    2 <NA>  1.34     c 2023-02-01         <NA> #VALUE! 23:01:02
#> 6  FALSE    2 <NA>  <NA> #NUM!       <NA>         <NA>       2 17:24:53
#> 7  FALSE    3 <NA>  1.56     e       <NA>         <NA>    <NA>     <NA>
#> 8  FALSE    1 <NA>   1.7     f 2023-03-02         <NA>     2.7 08:45:58
#> 9   <NA> <NA> <NA>  <NA>  <NA>       <NA>         <NA>    <NA>     <NA>
#> 10 FALSE    2 <NA>    23     h 2023-12-24         <NA>      25     <NA>
#> 11 FALSE    3 <NA>  67.3     i 2023-12-25         <NA>       3     <NA>
#> 12  <NA>    1 <NA>   123  <NA> 2023-07-31         <NA>     122     <NA>

# erase empty rows from dataset
wb_to_df(wb1, skip_empty_rows = TRUE)
#>     Var1 Var2 <NA>  Var3  Var4       Var5         Var6    Var7     Var8
#> 3   TRUE    1   NA     1     a 2023-05-29 3209324 This #DIV/0! 01:27:15
#> 4   TRUE   NA   NA #NUM!     b 2023-05-23         <NA>       0 14:02:57
#> 5   TRUE    2   NA  1.34     c 2023-02-01         <NA> #VALUE! 23:01:02
#> 6  FALSE    2   NA  <NA> #NUM!       <NA>         <NA>       2 17:24:53
#> 7  FALSE    3   NA  1.56     e       <NA>         <NA>    <NA>     <NA>
#> 8  FALSE    1   NA   1.7     f 2023-03-02         <NA>     2.7 08:45:58
#> 10 FALSE    2   NA    23     h 2023-12-24         <NA>      25     <NA>
#> 11 FALSE    3   NA  67.3     i 2023-12-25         <NA>       3     <NA>
#> 12    NA    1   NA   123  <NA> 2023-07-31         <NA>     122     <NA>

# erase empty columns from dataset
wb_to_df(wb1, skip_empty_cols = TRUE)
#>     Var1 Var2  Var3  Var4       Var5         Var6    Var7     Var8
#> 3   TRUE    1     1     a 2023-05-29 3209324 This #DIV/0! 01:27:15
#> 4   TRUE   NA #NUM!     b 2023-05-23         <NA>       0 14:02:57
#> 5   TRUE    2  1.34     c 2023-02-01         <NA> #VALUE! 23:01:02
#> 6  FALSE    2  <NA> #NUM!       <NA>         <NA>       2 17:24:53
#> 7  FALSE    3  1.56     e       <NA>         <NA>    <NA>     <NA>
#> 8  FALSE    1   1.7     f 2023-03-02         <NA>     2.7 08:45:58
#> 9     NA   NA  <NA>  <NA>       <NA>         <NA>    <NA>     <NA>
#> 10 FALSE    2    23     h 2023-12-24         <NA>      25     <NA>
#> 11 FALSE    3  67.3     i 2023-12-25         <NA>       3     <NA>
#> 12    NA    1   123  <NA> 2023-07-31         <NA>     122     <NA>

# convert first row to rownames
wb_to_df(wb1, sheet = 2, dims = "C6:G9", row_names = TRUE)
#>                mpg cyl disp  hp
#> Mazda RX4     21.0   6  160 110
#> Mazda RX4 Wag 21.0   6  160 110
#> Datsun 710    22.8   4  108  93

# define type of the data.frame
wb_to_df(wb1, cols = c(2, 5), types = c("Var1" = 0, "Var3" = 1))
#>     Var1   Var3
#> 3   TRUE   1.00
#> 4   TRUE    NaN
#> 5   TRUE   1.34
#> 6  FALSE     NA
#> 7  FALSE   1.56
#> 8  FALSE   1.70
#> 9   <NA>     NA
#> 10 FALSE  23.00
#> 11 FALSE  67.30
#> 12  <NA> 123.00

# start in row 5
wb_to_df(wb1, start_row = 5, col_names = FALSE)
#>        B  C  D      E     F          G  H       I        J
#> 5   TRUE  2 NA   1.34     c 2023-02-01 NA #VALUE! 23:01:02
#> 6  FALSE  2 NA     NA #NUM!       <NA> NA       2 17:24:53
#> 7  FALSE  3 NA   1.56     e       <NA> NA    <NA>     <NA>
#> 8  FALSE  1 NA   1.70     f 2023-03-02 NA     2.7 08:45:58
#> 9     NA NA NA     NA  <NA>       <NA> NA    <NA>     <NA>
#> 10 FALSE  2 NA  23.00     h 2023-12-24 NA      25     <NA>
#> 11 FALSE  3 NA  67.30     i 2023-12-25 NA       3     <NA>
#> 12    NA  1 NA 123.00  <NA> 2023-07-31 NA     122     <NA>

# na string
wb_to_df(wb1, na = "a")
#>     Var1 Var2 <NA>  Var3  Var4       Var5         Var6    Var7     Var8
#> 3   TRUE    1   NA     1  <NA> 2023-05-29 3209324 This #DIV/0! 01:27:15
#> 4   TRUE   NA   NA #NUM!     b 2023-05-23         <NA>       0 14:02:57
#> 5   TRUE    2   NA  1.34     c 2023-02-01         <NA> #VALUE! 23:01:02
#> 6  FALSE    2   NA  <NA> #NUM!       <NA>         <NA>       2 17:24:53
#> 7  FALSE    3   NA  1.56     e       <NA>         <NA>    <NA>     <NA>
#> 8  FALSE    1   NA   1.7     f 2023-03-02         <NA>     2.7 08:45:58
#> 9     NA   NA   NA  <NA>  <NA>       <NA>         <NA>    <NA>     <NA>
#> 10 FALSE    2   NA    23     h 2023-12-24         <NA>      25     <NA>
#> 11 FALSE    3   NA  67.3     i 2023-12-25         <NA>       3     <NA>
#> 12    NA    1   NA   123  <NA> 2023-07-31         <NA>     122     <NA>

# read names from row two and data starting from row 4
wb_to_df(wb1, dims = "B2:C2,B4:C+")
#>     Var1 Var2
#> 4   TRUE   NA
#> 5   TRUE    2
#> 6  FALSE    2
#> 7  FALSE    3
#> 8  FALSE    1
#> 9     NA   NA
#> 10 FALSE    2
#> 11 FALSE    3
#> 12    NA    1

###########################################################################
# Named regions
file_named_region <- system.file("extdata", "namedRegions3.xlsx", package = "openxlsx2")
wb2 <- wb_load(file_named_region)

# read dataset with named_region (returns global first)
wb_to_df(wb2, named_region = "MyRange", col_names = FALSE)
#>      A    B
#> 1 S2A1 S2B1

# read named_region from sheet
wb_to_df(wb2, named_region = "MyRange", sheet = 4, col_names = FALSE)
#>      A    B
#> 1 S3A1 S3B1

# read_xlsx() and wb_read()
example_file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
read_xlsx(file = example_file)
#>     Var1 Var2 <NA>  Var3  Var4       Var5         Var6    Var7     Var8
#> 3   TRUE    1   NA     1     a 2023-05-29 3209324 This #DIV/0! 01:27:15
#> 4   TRUE   NA   NA #NUM!     b 2023-05-23         <NA>       0 14:02:57
#> 5   TRUE    2   NA  1.34     c 2023-02-01         <NA> #VALUE! 23:01:02
#> 6  FALSE    2   NA  <NA> #NUM!       <NA>         <NA>       2 17:24:53
#> 7  FALSE    3   NA  1.56     e       <NA>         <NA>    <NA>     <NA>
#> 8  FALSE    1   NA   1.7     f 2023-03-02         <NA>     2.7 08:45:58
#> 9     NA   NA   NA  <NA>  <NA>       <NA>         <NA>    <NA>     <NA>
#> 10 FALSE    2   NA    23     h 2023-12-24         <NA>      25     <NA>
#> 11 FALSE    3   NA  67.3     i 2023-12-25         <NA>       3     <NA>
#> 12    NA    1   NA   123  <NA> 2023-07-31         <NA>     122     <NA>
df1 <- wb_read(file = example_file, sheet = 1)
df2 <- wb_read(file = example_file, sheet = 1, rows = c(1, 3, 5), cols = 1:3)
```
