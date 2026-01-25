# Create a data frame from a Workbook

The `wb_to_df()` function is the primary interface for extracting data
from spreadsheet files into R. It interprets the underlying XML
structure of a worksheet to reconstruct a data frame, handling cell
types, dimensions, and formatting according to user specification. While
`read_xlsx()` and `wb_read()` are available as streamlined internal
wrappers for users accustomed to other spreadsheet packages, wb_to_df()
serves as the foundational function and provides the most comprehensive
access to the package's data extraction and configuration parameters.

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

  A workbook file path, a
  [wbWorkbook](https://janmarvin.github.io/openxlsx2/dev/reference/wbWorkbook.md)
  object, or a valid URL.

- sheet:

  The name or index of the worksheet to read. Defaults to the first
  sheet.

- start_row, start_col:

  Optional numeric values specifying the first row or column to begin
  data discovery.

- row_names:

  Logical; if TRUE, uses the first column of the selection as row names.

- col_names:

  Logical; if TRUE, uses the first row of the selection as column
  headers.

- skip_empty_rows, skip_empty_cols:

  Logical; if TRUE, filters out rows or columns containing only missing
  values.

- skip_hidden_rows, skip_hidden_cols:

  Logical; if TRUE, excludes rows or columns marked as hidden in the
  worksheet metadata.

- rows, cols:

  Optional numeric vectors specifying the exact indices to read.

- detect_dates:

  Logical; if TRUE, identifies date and datetime styles for conversion.

- na:

  A character vector or a named list (e.g.,
  `list(strings = "", numbers = -99)`) defining values to treat as `NA`.

- fill_merged_cells:

  Logical; if TRUE, propagates the top-left value of a merged range to
  all cells in that range.

- dims:

  A character string defining the range. Supports wildcards (e.g.,
  "A1:++" or "A-:+5").

- show_formula:

  Logical; if TRUE, returns the formula strings instead of calculated
  values.

- convert:

  Logical; if TRUE, attempts to coerce columns to appropriate R classes.

- types:

  A named vector (numeric or character) to explicitly define column
  types.

- named_region:

  A character string referring to a defined name or spreadsheet Table.

- keep_attributes:

  Logical; if TRUE, attaches metadata such as the internal type table
  (tt) and types as attributes to the output.

- check_names:

  Logical; if TRUE, ensures column names are syntactically valid R names
  via [`make.names()`](https://rdrr.io/r/base/make.names.html).

- show_hyperlinks:

  Logical; if TRUE, replaces cell values with their underlying hyperlink
  targets.

- apply_numfmts:

  Logical; if TRUE, applies spreadsheet number formatting and returns
  strings.

- ...:

  Additional arguments passed to internal methods.

## Details

The function extracts data based on a defined range or the total data
extent of a worksheet. If `col_names = TRUE`, the first row of the
selection is treated as the header; otherwise, spreadsheet column
letters are used. If `row_names = TRUE`, the first column of the
selected range is assigned to the data frame's row names.

Dimension selection is highly flexible. The `dims` argument supports
standard "A1:B2" notation as well as dynamic wildcards for rows and
columns. Using `++` or `--` allows ranges to adapt to the spreadsheet's
content. For instance, `dims = "A2:C+"` reads from A2 to the last
available row in column C, while `dims = "A-:+9"` reads from the first
populated row in column A to the last column in row 9. If neither `dims`
nor `named_region` is provided, the function automatically calculates
the range based on the minimum and maximum populated cells, modified by
`start_row` and `start_col`.

Type conversion is governed by an internal guessing engine. If
`detect_dates` is enabled, serial dates are converted to R Date or
POSIXct objects. All datetimes are standardized to UTC. The function's
handling of time variables depends on the presence of the `hms` package;
if loaded, `wb_to_df()` returns `hms` variables. Otherwise, they are
returned as string variables in `hh:mm:ss` format. Users can provide
explicit column types via the `types` argument using numeric codes: 0
(character), 1 (numeric), 2 (Date), 3 (POSIXct), 4 (logical), 5 (hms),
and 6 (formula).

Regarding formulas, it is important to note that `wb_to_df()` will not
automatically evaluate formulas added to a workbook object via
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md).
In the underlying spreadsheet XML, only the formula expression is
written; the resulting value is typically generated by the spreadsheet
software's calculation engine when the file is opened and saved.
Consequently, reading a newly added formula cell without prior
evaluation in external software will result in an empty value unless
`show_formula = TRUE` is used to retrieve the formula string itself.

If `keep_attributes` is TRUE, the data frame is returned with additional
metadata. This includes the internal type-guessing table (`tt`), which
identifies the derived type for every cell in the range, and the
specific `types` vector used for conversion. These attributes are useful
for debugging or for applications requiring precise knowledge of the
spreadsheet's original cell metadata.

Specialized spreadsheet features include the ability to extract
hyperlink targets (`show_hyperlinks = TRUE`) instead of display text.
For complex layouts, `fill_merged_cells` propagates the value of a
top-left merged cell to all cells within the merge range. The `na`
argument supports sophisticated missing value definitions, accepting
either a character vector or a named list to differentiate between
string and numeric `NA` types.

## Notes

Recent versions of `openxlsx2` have introduced several changes to the
`wb_to_df()` API:

- Legacy arguments such as `na.strings` and `na.numbers` are no longer
  part of the public API and have been consolidated into the `na`
  argument.

- As of version 1.15, all datetime variables are imported with the
  timezone set to "UTC" to prevent system-specific local timezone
  shifts.

- The function now supports reverse-order or specific-order imports when
  a numeric vector is passed to the `rows` argument.

For extensive real-world examples and advanced usage patterns, consult
the package vignettes—specifically "openxlsx2 read to data frame"—and
the dedicated chapter in the `openxlsx2` book for real-life case
studies.

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
