# Helper to create a comment object

Creates a `wbComment` object. Use with
[`wb_add_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_add_comment.md)
to add to a worksheet location.

## Usage

``` r
wb_comment(
  text = NULL,
  style = NULL,
  visible = FALSE,
  author = getOption("openxlsx2.creator"),
  width = 2,
  height = 4
)
```

## Arguments

- text:

  Comment text. Character vector. or a
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md)
  string.

- style:

  A Style object or list of style objects the same length as comment
  vector.

- visible:

  Is comment visible? Default: `FALSE`.

- author:

  Author of comment. A string. By default, will look at
  `options("openxlsx2.creator")`. Otherwise, will check the system
  username.

- width:

  Textbox integer width in number of cells

- height:

  Textbox integer height in number of cells

## Value

A `wbComment` object

## Examples

``` r
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")

# write comment without author
c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
wb$add_comment(dims = "B10", comment = c1)

# Write another comment with author information
c2 <- wb_comment(text = "this is another comment", author = "Marco Polo")
wb$add_comment(sheet = 1, dims = "C10", comment = c2)

# write a styled comment with system author
s1 <- create_font(b = "true", color = wb_color(hex = "FFFF0000"), sz = "12")
s2 <- create_font(color = wb_color(hex = "FF000000"), sz = "9")
c3 <- wb_comment(text = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))

wb$add_comment(sheet = 1, dims = wb_dims(3, 6), comment = c3)
```
