# Internal comment functions

Users are advised to use
[`wb_add_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_comment.md)
and
[`wb_remove_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_comment.md).
`write_comment()` and `remove_comment()` are now deprecated. openxlsx2
will stop exporting it at some point in the future. Use the replacement
functions.

## Usage

``` r
write_comment(
  wb,
  sheet,
  col = NULL,
  row = NULL,
  comment,
  dims = rowcol_to_dim(row, col),
  color = NULL,
  file = NULL
)

remove_comment(
  wb,
  sheet,
  col = NULL,
  row = NULL,
  gridExpand = TRUE,
  dims = NULL
)
```

## Arguments

- wb:

  A workbook object

- sheet:

  A worksheet of the workbook

- row, col:

  Row and column of the cell

- comment:

  An object created by
  [`create_comment()`](https://janmarvin.github.io/openxlsx2/reference/create_comment.md)

- dims:

  Optional row and column as spreadsheet dimension, e.g. "A1"

- color:

  optional background color

- file:

  optional background image (file extension must be png or jpeg)

- gridExpand:

  If `TRUE`, all data in rectangle min(rows):max(rows) X
  min(cols):max(cols) will be removed.

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")
# add a comment without author
c1 <- wb_comment(text = "this is a comment", author = "")
wb$add_comment(dims = "B10", comment = c1)
#' # Remove comment
wb$remove_comment(sheet = "Sheet 1", dims = "B10")
# Write another comment with author information
c2 <- wb_comment(text = "this is another comment", author = "Marco Polo", visible = TRUE)
wb$add_comment(sheet = 1, dims = "C10", comment = c2)
# Works with formatted text also.
formatted_text <- fmt_txt("bar", underline = TRUE)
wb$add_comment(dims = "B5", comment = formatted_text)
# With background color
wb$add_comment(dims = "B7", comment = formatted_text, color = wb_color("green"))
# With background image. File extension must be png or jpeg, not jpg?
tmp <- tempfile(fileext = ".png")
png(file = tmp, bg = "transparent")
plot(1:10)
rect(1, 5, 3, 7, col = "white")
dev.off()
#> agg_record_147083781 
#>                    2 

c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
wb$add_comment(dims = "B12", comment = c1, file = tmp)
```
