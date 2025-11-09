# Add data to a worksheet

Add data to worksheet with optional styling.

## Usage

``` r
wb_add_data(
  wb,
  sheet = current_sheet(),
  x,
  dims = wb_dims(start_row, start_col),
  start_col = 1,
  start_row = 1,
  array = FALSE,
  col_names = TRUE,
  row_names = FALSE,
  with_filter = FALSE,
  name = NULL,
  sep = ", ",
  apply_cell_style = TRUE,
  remove_cell_style = FALSE,
  na.strings = na_strings(),
  inline_strings = TRUE,
  enforce = FALSE,
  ...
)
```

## Arguments

- wb:

  A Workbook object containing a worksheet.

- sheet:

  The worksheet to write to. Can be the worksheet index or name.

- x:

  Object to be written. For classes supported look at the examples.

- dims:

  Spreadsheet cell range that will determine `start_col` and
  `start_row`: "A1", "A1:B2", "A:B"

- start_col:

  A vector specifying the starting column to write `x` to.

- start_row:

  A vector specifying the starting row to write `x` to.

- array:

  A bool if the function written is of type array

- col_names:

  If `TRUE`, column names of `x` are written.

- row_names:

  If `TRUE`, the row names of `x` are written.

- with_filter:

  If `TRUE`, add filters to the column name row. NOTE: can only have one
  filter per worksheet.

- name:

  The name of a named region if specified.

- sep:

  Only applies to list columns. The separator used to collapse list
  columns to a character vector e.g.
  `sapply(x$list_column, paste, collapse = sep)`.

- apply_cell_style:

  Should we write cell styles to the workbook

- remove_cell_style:

  keep the cell style?

- na.strings:

  Value used for replacing `NA` values from `x`. Default looks if
  `options(openxlsx2.na.strings)` is set. Otherwise
  [`na_strings()`](https://janmarvin.github.io/openxlsx2/dev/reference/waivers.md)
  uses the special `#N/A` value within the workbook.

- inline_strings:

  write characters as inline strings

- enforce:

  enforce that selected dims is filled. For this to work, `dims` must
  match `x`

- ...:

  additional arguments

## Value

A `wbWorkbook`, invisibly.

## Details

Formulae written using
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md)
to a Workbook object will not get picked up by
[`read_xlsx()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_to_df.md).
This is because only the formula is written into the worksheet and it
will be evaluated once the file is opened in spreadsheet software. The
string `"_openxlsx_NA"` is reserved for `openxlsx2`. If the data frame
contains this string, the output will be broken. Similar factor labels
`"_openxlsx_Inf"`, `"_openxlsx_nInf"`, and `"_openxlsx_NaN"` are
reserved.

Supported classes are data frames, matrices and vectors of various types
and everything that can be converted into a data frame with
[`as.data.frame()`](https://rdrr.io/r/base/as.data.frame.html).
Everything else that the user wants to write should either be converted
into a vector or data frame or written in vector or data frame segments.
This includes base classes such as `table`, which were coerced
internally in the predecessor of this package.

Even vectors and data frames can consist of different classes. Many base
classes are covered, though not all and far from all third-party
classes. When data of an unknown class is written, it is handled with
[`as.character()`](https://rdrr.io/r/base/character.html). It is not
possible to write character nodes beginning with `<r>` or `<r/>`. Both
are reserved for internal functions. If you need these. You have to wrap
the input string in
[`fmt_txt()`](https://janmarvin.github.io/openxlsx2/dev/reference/fmt_txt.md).

The columns of `x` with class Date/POSIXt, currency, accounting,
hyperlink, percentage are automatically styled as dates, currency,
accounting, hyperlinks, percentages respectively. When writing POSIXt,
the users local timezone should not matter. The openxml standard does
not have a timezone and the conversion from the local timezone should
happen internally, so that date and time are converted, but the timezone
is dropped. This conversion could cause a minor precision loss. The
datetime in R and in spreadsheets might differ by 1 second, caused by
floating point precision. When read from the worksheet, starting with
`openxlsx2` release `1.15` the datetime is returned in `"UTC"`.

Functions `wb_add_data()` and
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md)
behave quite similar. The distinction is that the latter creates a table
in the worksheet that can be used for different kind of formulas and can
be sorted independently, though is less flexible than basic cell
regions.

## See also

Other workbook wrappers:
[`base_font-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/base_font-wb.md),
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`creators-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/creators-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_chartsheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_chartsheet.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_worksheet.md),
[`wb_base_colors`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_base_colors.md),
[`wb_clone_worksheet()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_clone_worksheet.md),
[`wb_copy_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_copy_cells.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md),
[`wb_save()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_save.md),
[`wb_set_last_modified_by()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_set_last_modified_by.md),
[`wb_workbook()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_workbook.md)

Other worksheet content functions:
[`col_widths-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/col_widths-wb.md),
[`filter-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/filter-wb.md),
[`grouping-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/grouping-wb.md),
[`named_region-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/named_region-wb.md),
[`row_heights-wb`](https://janmarvin.github.io/openxlsx2/dev/reference/row_heights-wb.md),
[`wb_add_conditional_formatting()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_conditional_formatting.md),
[`wb_add_data_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_data_table.md),
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md),
[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md),
[`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_pivot_table.md),
[`wb_add_slicer()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_slicer.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_thread.md),
[`wb_freeze_pane()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_freeze_pane.md),
[`wb_merge_cells()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_merge_cells.md)

## Examples

``` r
## See formatting vignette for further examples.

## Options for default styling (These are the defaults)
options("openxlsx2.dateFormat" = "mm/dd/yyyy")
options("openxlsx2.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
options("openxlsx2.numFmt" = NULL)

#############################################################################
## Create Workbook object and add worksheets
wb <- wb_workbook()

## Add worksheets
wb$add_worksheet("Cars")
wb$add_worksheet("Formula")

x <- mtcars[1:6, ]
wb$add_data("Cars", x, start_col = 2, start_row = 3, row_names = TRUE)

#############################################################################
## Hyperlinks
## - vectors/columns with class 'hyperlink' are written as hyperlinks'

v <- rep("https://CRAN.R-project.org/", 4)
names(v) <- paste0("Hyperlink", 1:4) # Optional: names will be used as display text
class(v) <- "hyperlink"
wb$add_data("Cars", x = v, dims = "B32")

#############################################################################
## Formulas
## - vectors/columns with class 'formula' are written as formulas'

df <- data.frame(
  x = 1:3, y = 1:3,
  z = paste(paste0("A", 1:3 + 1L), paste0("B", 1:3 + 1L), sep = "+"),
  stringsAsFactors = FALSE
)

class(df$z) <- c(class(df$z), "formula")

wb$add_data(sheet = "Formula", x = df)

#############################################################################
# update cell range and add mtcars
xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
wb2 <- wb_load(xlsxFile)

# read dataset with inlinestr
wb_to_df(wb2)
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
wb2 <- wb_add_data(wb2, sheet = 1, mtcars, dims = wb_dims(4, 4))
wb_to_df(wb2)
#>     Var1 Var2 <NA> Var3  Var4       Var5         Var6    Var7     Var8 <NA>
#> 3   TRUE    1 <NA>    1     a 2023-05-29 3209324 This #DIV/0! 01:27:15 <NA>
#> 4   TRUE   NA  mpg  cyl  disp         hp         drat      wt     qsec   vs
#> 5   TRUE    2   21    6   160 1900-04-19          3.9    2.62 11:02:24    0
#> 6  FALSE    2   21    6   160        110          3.9   2.875 00:28:48    0
#> 7  FALSE    3 22.8    4   108         93         3.85    2.32    18.61    1
#> 8  FALSE    1 21.4    6   258 1900-04-19         3.08   3.215 10:33:36    1
#> 9     NA   NA 18.7    8   360        175         3.15    3.44    17.02    0
#> 10 FALSE    2 18.1    6   225 1900-04-14         2.76    3.46    20.22    1
#> 11 FALSE    3 14.3    8   360 1900-09-01         3.21    3.57    15.84    0
#> 12    NA    1 24.4    4 146.7 1900-03-02         3.69    3.19       20    1
#> 13    NA   NA 22.8    4 140.8         95         3.92    3.15     22.9    1
#> 14    NA   NA 19.2    6 167.6        123         3.92    3.44     18.3    1
#> 15    NA   NA 17.8    6 167.6        123         3.92    3.44     18.9    1
#> 16    NA   NA 16.4    8 275.8        180         3.07    4.07     17.4    0
#> 17    NA   NA 17.3    8 275.8        180         3.07    3.73     17.6    0
#> 18    NA   NA 15.2    8 275.8        180         3.07    3.78       18    0
#> 19    NA   NA 10.4    8   472        205         2.93    5.25    17.98    0
#> 20    NA   NA 10.4    8   460        215            3   5.424    17.82    0
#> 21    NA   NA 14.7    8   440        230         3.23   5.345    17.42    0
#> 22    NA   NA 32.4    4  78.7         66         4.08     2.2    19.47    1
#> 23    NA   NA 30.4    4  75.7         52         4.93   1.615    18.52    1
#> 24    NA   NA 33.9    4  71.1         65         4.22   1.835     19.9    1
#> 25    NA   NA 21.5    4 120.1         97          3.7   2.465    20.01    1
#> 26    NA   NA 15.5    8   318        150         2.76    3.52    16.87    0
#> 27    NA   NA 15.2    8   304        150         3.15   3.435     17.3    0
#> 28    NA   NA 13.3    8   350        245         3.73    3.84    15.41    0
#> 29    NA   NA 19.2    8   400        175         3.08   3.845    17.05    0
#> 30    NA   NA 27.3    4    79         66         4.08   1.935     18.9    1
#> 31    NA   NA   26    4 120.3         91         4.43    2.14     16.7    0
#> 32    NA   NA 30.4    4  95.1        113         3.77   1.513     16.9    1
#> 33    NA   NA 15.8    8   351        264         4.22    3.17     14.5    0
#> 34    NA   NA 19.7    6   145        175         3.62    2.77     15.5    0
#> 35    NA   NA   15    8   301        335         3.54    3.57     14.6    0
#> 36    NA   NA 21.4    4   121        109         4.11    2.78     18.6    1
#>    <NA> <NA> <NA>
#> 3  <NA> <NA> <NA>
#> 4    am gear carb
#> 5     1    4    4
#> 6     1    4    4
#> 7     1    4    1
#> 8     0    3    1
#> 9     0    3    2
#> 10    0    3    1
#> 11    0    3    4
#> 12    0    4    2
#> 13    0    4    2
#> 14    0    4    4
#> 15    0    4    4
#> 16    0    3    3
#> 17    0    3    3
#> 18    0    3    3
#> 19    0    3    4
#> 20    0    3    4
#> 21    0    3    4
#> 22    1    4    1
#> 23    1    4    2
#> 24    1    4    1
#> 25    0    3    1
#> 26    0    3    2
#> 27    0    3    2
#> 28    0    3    4
#> 29    0    3    2
#> 30    1    4    1
#> 31    1    5    2
#> 32    1    5    2
#> 33    1    5    4
#> 34    1    5    6
#> 35    1    5    8
#> 36    1    4    2
```
