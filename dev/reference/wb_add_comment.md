# Add comment to worksheet

Add comment to worksheet

## Usage

``` r
wb_add_comment(wb, sheet = current_sheet(), dims = "A1", comment, ...)

wb_get_comment(wb, sheet = current_sheet(), dims = NULL)

wb_remove_comment(wb, sheet = current_sheet(), dims = "A1", ...)
```

## Arguments

- wb:

  A workbook object

- sheet:

  A worksheet of the workbook

- dims:

  Optional row and column as spreadsheet dimension, e.g. "A1"

- comment:

  A comment to apply to `dims` created by
  [`wb_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_comment.md),
  a string or a
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/dev/reference/fmt_txt.md)
  object

- ...:

  additional arguments

## Value

The Workbook object, invisibly.

## Details

If applying a `comment` with a string, it will use
[`wb_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_comment.md)
default values. If additional background colors are applied, RGB colors
should be provided, either as hex code or with builtin R colors. The
alpha channel is ignored.

## See also

[`wb_comment()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_comment.md),
[`wb_add_thread()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_add_thread.md)

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
#> agg_record_120551243 
#>                    2 

c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
wb$add_comment(dims = "B12", comment = c1, file = tmp)
```
