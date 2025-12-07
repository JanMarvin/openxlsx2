# xlsx reading, writing and editing.

This R package is a modern reinterpretation of the widely used popular
`openxlsx` package. Similar to its predecessor, it simplifies the
creation of xlsx files by providing a clean interface for writing,
designing and editing worksheets. Based on a powerful XML library and
focusing on modern programming flows in pipes or chains, `openxlsx2`
allows to break many new ground.

## Details

The `openxlsx2` package provides comprehensive functionality for
interacting with Office Open XML spreadsheet files. Users can read data
using
[`read_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/wb_to_df.md)
and write data to spreadsheets via
[`write_xlsx()`](https://janmarvin.github.io/openxlsx2/reference/write_xlsx.md),
with options to specify sheet names and cell ranges for targeted
operations. Beyond basic read/write capabilities, `openxlsx2`
facilitates extensive workbook
([`wb_workbook()`](https://janmarvin.github.io/openxlsx2/reference/wb_workbook.md))
manipulations, including:

- Loading a workbook into R with
  [`wb_load()`](https://janmarvin.github.io/openxlsx2/reference/wb_load.md)
  and saving it with
  [`wb_save()`](https://janmarvin.github.io/openxlsx2/reference/wb_save.md)

- Adding/removing and modifying worksheets and data with
  [`wb_add_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_worksheet.md),
  [`wb_remove_worksheet()`](https://janmarvin.github.io/openxlsx2/reference/wb_remove_worksheet.md),
  and
  [`wb_add_data()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data.md).

- Enhancing spreadsheets with comments
  ([`wb_add_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_comment.md)),
  images
  ([`wb_add_image()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_image.md)),
  plots
  ([`wb_add_plot()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_plot.md)),
  charts
  ([`wb_add_mschart()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_mschart.md)),
  and pivot tables
  ([`wb_add_pivot_table()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_pivot_table.md)).
  Customizing cell styles using fonts
  ([`wb_add_font()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_font.md)),
  number formats
  ([`wb_add_numfmt()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_numfmt.md)),
  backgrounds
  ([`wb_add_fill()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_fill.md)),
  and alignments
  ([`wb_add_cell_style()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_cell_style.md)).
  Inserting custom text strings with
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md)
  and creating comprehensive table styles with
  [`create_tablestyle()`](https://janmarvin.github.io/openxlsx2/reference/create_tablestyle.md).

### Interaction

Interaction with `openxlsx2` objects can occur through two primary
methods:

*Wrapper Function Method*: Utilizes the `wb` family of functions that
support piping to streamline operations.

    wb <- wb_workbook(creator = "My name here") |>
      wb_add_worksheet(sheet = "Expenditure", grid_lines = FALSE) |>
      wb_add_data(x = USPersonalExpenditure, row_names = TRUE)

*Chaining Method*: Directly modifies the object through a series of
chained function calls.

    wb <- wb_workbook(creator = "My name here")$
      add_worksheet(sheet = "Expenditure", grid_lines = FALSE)$
      add_data(x = USPersonalExpenditure, row_names = TRUE)

While wrapper functions require explicit assignment of their output to
reflect changes, chained functions inherently modify the input object.
Both approaches are equally supported, offering flexibility to suit user
preferences. The documentation mainly highlights the use of wrapper
functions. To find information, users should look up the wb function
name e.g.
[`?wb_add_data_table`](https://janmarvin.github.io/openxlsx2/reference/wb_add_data_table.md)
rather than searching for
[`?wbWorkbook`](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md).

Function arguments follow the snake_case convention, but for backward
compatibility, camelCase is also supported at the moment. The API aims
to maintain consistency in its arguments, with a special focus on
`sheet`
([`wb_get_sheet_names()`](https://janmarvin.github.io/openxlsx2/reference/sheet_names-wb.md))
and `dims`
([wb_dims](https://janmarvin.github.io/openxlsx2/reference/wb_dims.md)),
which are of particular importance to users.

### Locale

By default, `openxlsx2` uses the American English word for color
(written with 'o' instead of the British English 'ou'). However, both
spellings are supported. So where the documentation uses a 'color', the
function should also accept a 'colour'. However, this is not indicated
by the autocompletion.

### Numeric Precision

R typically uses IEEE 754 double-precision floating-point numbers. When
`openxlsx2` writes these R numeric values to a spreadsheet, it stores
them using double precision (and shortens if possible) within the file
structure. However, spreadsheet software operates with, displays, and
performs calculations using only about 15 significant decimal digits. If
spreadsheet software encounters a number with higher precision (like one
written by `openxlsx2`), it will typically round that number to 15
significant digits for its internal use and display. Conversely, when
reading numeric data from an xlsx/xlsm file, `openxlsx2` reads the
stored double-precision value. However, potential discrepancies can
still arise. If the number was originally created or calculated within
spreadsheet software and exceeded 15 digits, it may have already rounded
before saving. Additionally, subtle rounding differences can sometimes
occur during the XML-to-numeric conversion when reading the file into R.
Expect minor variations, especially in the least significant digits, and
avoid direct equality (`==`) comparisons; check for differences within a
small tolerance instead.

### Supported files

Supported input files include `xlsx`, `xlsm`, and `xlsb`. The `xlsx` and
`xlsm` formats are fully supported. The key difference between these two
formats is that `xlsm` files may contain a binary blob that stores VBA
code, while `xlsx` files do not.

Support for the `xlsb` format is more limited. A custom parser is used
to convert the binary format into a pseudo-XML structure that can be
loaded into a
[`wbWorkbook`](https://janmarvin.github.io/openxlsx2/reference/wbWorkbook.md).
This allows `xlsb` files to be handled similarly to other workbook
formats. However, the parser does not fully implement the entire `xlsb`
specification. It provides functionality to read worksheets into data
frames and to save the content as an `xlsx` or `xlsm` file for
comparison of values and formulas. Some components, such as pivot tables
and conditional formatting, are not currently parsed. Writing `xlsb`
files is not supported, and big endian support is very limited. At
present, there are no plans to extend support for `xlsb` files further,
due to the complexity of the format. It is recommended to use
spreadsheet software to convert files to or from the `xlsb` format when
necessary.

The XML-based formats (`xlsx` and `xlsm`) are fully supported. Loading,
modifying, and saving these files with minimal unintended changes is a
core project goal. Nevertheless, the XML format can be fragile.
Modifying cells that interact with formulas, named objects such as
tables or regions, slicers, or pivot tables can lead to unexpected
behavior. It is therefore strongly advised to keep backups of important
files and to regularly verify the output using appropriate spreadsheet
software.

### Authors and contributions

For a full list of all authors that have made this package possible and
for whom we are grateful, please see:

    system.file("AUTHORS", package = "openxlsx2")

If you feel like you should be included on this list, please let us
know. If you have something to contribute, you are welcome. If something
is not working as expected, open issues or if you have solved an issue,
open a pull request. Please be respectful and be aware that we are
volunteers doing this for fun in our unpaid free time. We will work on
problems when we have time or need.

### License

This package is licensed under the MIT license and is based on
[`openxlsx`](https://github.com/ycphs/openxlsx) (by Alexander Walker and
Philipp Schauberger; COPYRIGHT 2014-2022) and
[`pugixml`](https://github.com/zeux/pugixml) (by Arseny Kapoulkine;
COPYRIGHT 2006-2025). Both released under the MIT license.

## See also

- `browseVignettes("openxlsx2")`

- <https://janmarvin.github.io/openxlsx2/>

- <https://janmarvin.github.io/ox2-book/> for examples

## Author

**Maintainer**: Jan Marvin Garbuszus <jan.garbuszus@ruhr-uni-bochum.de>

Authors:

- Jordan Mark Barbone <jmbarbone@gmail.com>
  ([ORCID](https://orcid.org/0000-0001-9788-3628))

Other contributors:

- Olivier Roy \[contributor\]

- openxlsx authors (openxlsx package) \[copyright holder\]

- Arseny Kapoulkine (Author of included pugixml code) \[contributor,
  copyright holder\]

## Examples

``` r
# read xlsx or xlsm files
path <- system.file("extdata/openxlsx2_example.xlsx", package = "openxlsx2")
read_xlsx(path)
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

# or import workbooks
wb <- wb_load(path)

# read a data frame
wb_to_df(wb)
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
