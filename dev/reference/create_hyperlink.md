# Create spreadsheet hyperlink string

Wrapper to create internal hyperlink string to pass to
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md).
Either link to external URLs or local files or straight to cells of
local xlsx sheets.

Note that for an external URL, only `file` and `text` should be
supplied. You can supply `dims` to
[`wb_add_formula()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_formula.md)
to control the location of the link.

## Usage

``` r
create_hyperlink(sheet, row = 1, col = 1, text = NULL, file = NULL)
```

## Arguments

- sheet:

  Name of a worksheet

- row:

  integer row number for hyperlink to link to

- col:

  column number of letter for hyperlink to link to

- text:

  Display text

- file:

  Hyperlink or xlsx file name to point to. If `NULL`, hyperlink is
  internal.

## See also

[`wb_add_hyperlink()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_hyperlink.md)

## Examples

``` r
wb <- wb_workbook()$
  add_worksheet("Sheet1")$add_worksheet("Sheet2")$add_worksheet("Sheet3")

## Internal Hyperlink - create hyperlink formula manually
x <- '=HYPERLINK(\"#Sheet2!B3\", "Text to Display - Link to Sheet2")'
wb$add_formula(sheet = "Sheet1", x = x, dims = "A1")

## Internal - No text to display using create_hyperlink() function
x <- create_hyperlink(sheet = "Sheet3", row = 1, col = 2)
wb$add_formula(sheet = "Sheet1", x = x, dims = "A2")

## Internal - Text to display
x <- create_hyperlink(sheet = "Sheet3", row = 1, col = 2,text = "Link to Sheet 3")
wb$add_formula(sheet = "Sheet1", x = x, dims = "A3")

## Link to file - No text to display
fl <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
x <- create_hyperlink(sheet = "Sheet1", row = 3, col = 10, file = fl)
wb$add_formula(sheet = "Sheet1", x = x, dims = "A4")

## Link to file - Text to display
fl <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
x <- create_hyperlink(sheet = "Sheet2", row = 3, col = 10, file = fl, text = "Link to File.")
wb$add_formula(sheet = "Sheet1", x = x, dims = "A5")

## Link to external file - Text to display
x <- '=HYPERLINK("[C:/Users]", "Link to an external file")'
wb$add_formula(sheet = "Sheet1", x = x, dims = "A6")

x <- create_hyperlink(text = "test.png", file = "D:/somepath/somepicture.png")
wb$add_formula(x = x, dims = "A7")


## Link to an URL.
x <- create_hyperlink(text = "openxlsx2 website", file = "https://janmarvin.github.io/openxlsx2/")

wb$add_formula(x = x, dims = "A8")
# if (interactive()) wb$open()
```
