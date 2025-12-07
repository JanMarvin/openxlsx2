# Create a comment

Use
[`wb_comment()`](https://janmarvin.github.io/openxlsx2/reference/wb_comment.md)
in new code. See
[openxlsx2-deprecated](https://janmarvin.github.io/openxlsx2/reference/openxlsx2-deprecated.md)

## Usage

``` r
create_comment(
  text,
  author = Sys.info()[["user"]],
  style = NULL,
  visible = TRUE,
  width = 2,
  height = 4
)
```

## Arguments

- text:

  Comment text. Character vector. or a
  [`fmt_txt()`](https://janmarvin.github.io/openxlsx2/reference/fmt_txt.md)
  string.

- author:

  A string, by default, will use "user"

- style:

  A Style object or list of style objects the same length as comment
  vector.

- visible:

  Default: `TRUE`. Is the comment visible by default?

- width:

  Textbox integer width in number of cells

- height:

  Textbox integer height in number of cells

## Value

a `wbComment` object
