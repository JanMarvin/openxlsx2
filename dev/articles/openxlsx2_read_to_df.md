# openxlsx2 read to data frame

## Importing data

Coming from `openxlsx` you might know about `read.xlsx()` (two
functions, one for files and one for workbooks) and `readWorkbook()`.
Functions that do different things, but mostly the same. In `openxlsx2`
we tried our best to reduce the complexity under the hood and for the
user as well. In `openxlsx2` they are replaced with
[`read_xlsx()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_to_df.md),
[`wb_read()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_to_df.md)
and they share the same underlying function
[`wb_to_df()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_to_df.md).

For this example we will use example data provided by the package. You
can locate it in our “inst/extdata” folder. The files are included with
the package source and you can open them in any calculation software as
well.

### Basic import

We begin with the `openxlsx2_example.xlsx` file by telling R where to
find this file on our system

``` r
file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
```

The object contains a path to the xlsx file and we pass this file to our
function to read the workbook into R

``` r
# import workbook
library(openxlsx2)
wb_to_df(file)
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
```

The output is created as a data frame and contains data types date,
logical, numeric and character. The function to import the file to R,
[`wb_to_df()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_to_df.md)
provides similar options as the `openxlsx` functions `read.xlsx()` and
`readWorkbook()` and a few new functions we will go through the options.
As you might have noticed, we return the column of the xlsx file as the
row name of the data frame returned. Per default the first sheet in the
workbook is imported. If you want to switch this, either provide the
`sheet` parameter with the correct index or provide the sheet name.

### `col_names` - first row as column name

In the previous example the first imported row was used as column name
for the data frame. This is the default behavior, but not always wanted
or expected. Therefore this behavior can be disabled by the user.

``` r
# do not convert first row to column names
wb_to_df(file, col_names = FALSE)
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
```

### `detect_dates` - convert cells to R dates

The creators of the openxml standard are well known for mistakenly
treating something as a date and `openxlsx2` has built in ways to
identify a cell as a date and will try to convert the value for you, but
unfortunately this is not always a trivial task and might fail. In such
a case we provide an option to disable the date conversion entirely. In
this case the underlying numerical value will be returned.

``` r
# do not try to identify dates in the data
wb_to_df(file, detect_dates = FALSE)
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
```

### `show_formula` - show formulas instead of results

Sometimes things might feel off. This can be because the openxml files
are not updating formula results in the sheets unless they are opened in
software that provides such functionality as certain tabular calculation
software. Therefore the user might be interested in the underlying
functions to see what is going on in the sheet. Using `show_formula`
this is possible

``` r
# return the underlying spreadsheet formula instead of their values
wb_to_df(file, show_formula = TRUE)
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
```

### `dims` - read specific dimension

Sometimes the entire worksheet contains to much data, in such case we
provide functions to read only a selected dimension range. Such a range
consists of either a specific cell like “A1” or a cell range in the
notion used in the `openxml` standard

``` r
# read dimension without column names
wb_to_df(file, dims = "A2:C5", col_names = FALSE)
#>    A    B    C
#> 2 NA Var1 Var2
#> 3 NA TRUE    1
#> 4 NA TRUE <NA>
#> 5 NA TRUE    2
```

Alternatively, if you don’t know the spreadsheet’s A1 notation, you can
use
[`wb_dims()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_dims.md)
to specify the dimension. See below or
in[`?wb_dims`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_dims.md)
for more details.

``` r
# read dimension without column names with `wb_dims()`
wb_to_df(file, dims = wb_dims(rows = 2:5, cols = 1:3), col_names = FALSE)
#>    A    B    C
#> 2 NA Var1 Var2
#> 3 NA TRUE    1
#> 4 NA TRUE <NA>
#> 5 NA TRUE    2
```

### `cols` - read selected columns

If you do not want to read a specific cell, but a cell range you can use
the column attribute. This attribute takes a numeric vector as argument

``` r
# read selected cols
wb_to_df(file, cols = c("A:B", "G"))
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
```

### `rows` - read selected rows

The same goes with rows. You can select them using numeric vectors

``` r
# read selected rows
wb_to_df(file, rows = c(2, 4, 6))
#>    Var1 Var2 <NA>  Var3  Var4       Var5 Var6 Var7     Var8
#> 4  TRUE   NA   NA #NUM!     b 2023-05-23   NA    0 14:02:57
#> 6 FALSE    2   NA  <NA> #NUM!       <NA>   NA    2 17:24:53
```

### `convert` - convert input to guessed type

In xml exists no difference between value types. All values are per
default characters. To provide these as numerics, logicals or dates,
`openxlsx2` and every other software dealing with xlsx files has to make
assumptions about the cell type. This is especially tricky due to the
notion of worksheets. Unlike in a data frame, a worksheet can have a
wild mix of all types of data. Even though the conversion process from
character to date or numeric is rather solid, sometimes the user might
want to see the data without any conversion applied. This might be
useful in cases where something unexpected happened or the import
created warnings. In such a case you can look at the raw input data. If
you want to disable date detection as well, please see the entry above.

``` r
# convert characters to numerics and date (logical too?)
wb_to_df(file, convert = FALSE)
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
```

### `skip_empty_rows` - remove empty rows

Even though `openxlsx2` imports everything as requested, sometimes it
might be helpful to remove empty lines from the data. These might be
either left empty intentional or empty because they are were formatted,
but the cell value was removed afterwards. This was added mostly for
backward comparability, but the default has been changed to `FALSE`. The
behavior has changed a bit as well. Previously empty cells were removed
prior to the conversion to R data frames, now they are removed after the
conversion and are removed only if they are completely empty

``` r
# erase empty rows from dataset
tail(wb_to_df(file, sheet = 1, skip_empty_rows = TRUE))
#>     Var1 Var2 <NA> Var3  Var4       Var5 Var6 Var7     Var8
#> 6  FALSE    2   NA <NA> #NUM!       <NA> <NA>    2 17:24:53
#> 7  FALSE    3   NA 1.56     e       <NA> <NA> <NA>     <NA>
#> 8  FALSE    1   NA  1.7     f 2023-03-02 <NA>  2.7 08:45:58
#> 10 FALSE    2   NA   23     h 2023-12-24 <NA>   25     <NA>
#> 11 FALSE    3   NA 67.3     i 2023-12-25 <NA>    3     <NA>
#> 12    NA    1   NA  123  <NA> 2023-07-31 <NA>  122     <NA>
```

### `skip_empty_cols` - remove empty columns

The same for columns

``` r
# erase empty cols from dataset
wb_to_df(file, skip_empty_cols = TRUE)
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
```

### `row_names` - keep rownames from input

Sometimes the data source might provide rownames as well. In such a case
you can `openxlsx2` to treat the first column as rowname

``` r
# convert first row to rownames
wb_to_df(file, sheet = 2, dims = "C6:G9", row_names = TRUE)
#>                mpg cyl disp  hp
#> Mazda RX4     21.0   6  160 110
#> Mazda RX4 Wag 21.0   6  160 110
#> Datsun 710    22.8   4  108  93
```

### `types` - convert column to specific type

If the user know better than the software what type to expect in a
worksheet, this can be provided via types. This parameter takes a named
numeric. `0` is character, `1` is numeric and `2` is date

``` r
# define type of the data.frame
wb_to_df(file, cols = c(2, 5), types = c("Var1" = 0, "Var3" = 1))
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
```

### `start_row` - where to begin

Often the creator of the worksheet has used a lot of creativity and the
data does not begin in the first row, instead it begins somewhere else.
To define the row where to begin reading, define it via the `start_row`
parameter

``` r
# start in row 5
wb_to_df(file, start_row = 5, col_names = FALSE)
#>        B  C  D      E     F          G  H       I        J
#> 5   TRUE  2 NA   1.34     c 2023-02-01 NA #VALUE! 23:01:02
#> 6  FALSE  2 NA     NA #NUM!       <NA> NA       2 17:24:53
#> 7  FALSE  3 NA   1.56     e       <NA> NA    <NA>     <NA>
#> 8  FALSE  1 NA   1.70     f 2023-03-02 NA     2.7 08:45:58
#> 9     NA NA NA     NA  <NA>       <NA> NA    <NA>     <NA>
#> 10 FALSE  2 NA  23.00     h 2023-12-24 NA      25     <NA>
#> 11 FALSE  3 NA  67.30     i 2023-12-25 NA       3     <NA>
#> 12    NA  1 NA 123.00  <NA> 2023-07-31 NA     122     <NA>
```

### `na.strings` - define missing values

There is the “#N/A” string, but often the user will be faced with custom
missing values and other values we are not interested. Such strings can
be passed as character vector via `na.strings`

``` r
# na strings
wb_to_df(file, na.strings = "")
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
```
