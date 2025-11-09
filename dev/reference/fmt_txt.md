# format strings independent of the cell style.

format strings independent of the cell style.

## Usage

``` r
fmt_txt(
  x,
  bold = FALSE,
  italic = FALSE,
  underline = FALSE,
  strike = FALSE,
  size = NULL,
  color = NULL,
  font = NULL,
  charset = NULL,
  outline = NULL,
  vert_align = NULL,
  family = NULL,
  shadow = NULL,
  condense = NULL,
  extend = NULL,
  ...
)

# S3 method for class 'fmt_txt'
x + y

# S3 method for class 'fmt_txt'
as.character(x, ...)

# S3 method for class 'fmt_txt'
print(x, ...)
```

## Arguments

- x, y:

  an openxlsx2 fmt_txt string

- bold, italic, underline, strike:

  logical defaulting to `FALSE`

- size:

  the font size

- color:

  a
  [`wb_color()`](https://janmarvin.github.io/openxlsx2/dev/reference/wb_color.md)
  for the font

- font:

  the font name

- charset:

  integer value from the table below

- outline, shadow, condense, extend:

  logical defaulting to `NULL`

- vert_align:

  any of `baseline`, `superscript`, or `subscript`

- family:

  a font family

- ...:

  additional arguments

## Details

The result is an xml string. It is possible to paste multiple
`fmt_txt()` strings together to create a string with differing styles.
It is possible to supply different underline styles to `underline`.

Using `fmt_txt(charset = 161)` will give the Greek Character Set

|         |                       |
|---------|-----------------------|
| charset | "Character Set"       |
| 0       | "ANSI_CHARSET"        |
| 1       | "DEFAULT_CHARSET"     |
| 2       | "SYMBOL_CHARSET"      |
| 77      | "MAC_CHARSET"         |
| 128     | "SHIFTJIS_CHARSET"    |
| 129     | "HANGUL_CHARSET"      |
| 130     | "JOHAB_CHARSET"       |
| 134     | "GB2312_CHARSET"      |
| 136     | "CHINESEBIG5_CHARSET" |
| 161     | "GREEK_CHARSET"       |
| 162     | "TURKISH_CHARSET"     |
| 163     | "VIETNAMESE_CHARSET"  |
| 177     | "HEBREW_CHARSET"      |
| 178     | "ARABIC_CHARSET"      |
| 186     | "BALTIC_CHARSET"      |
| 204     | "RUSSIAN_CHARSET"     |
| 222     | "THAI_CHARSET"        |
| 238     | "EASTEUROPE_CHARSET"  |
| 255     | "OEM_CHARSET"         |

You can join additional objects into fmt_txt() objects using "+". Though
be aware that `fmt_txt("sum:") + (2 + 2)` is different to
`fmt_txt("sum:") + 2 + 2`.

## See also

[`create_font()`](https://janmarvin.github.io/openxlsx2/dev/reference/create_font.md)

## Examples

``` r
fmt_txt("bar", underline = TRUE)
#> fmt_txt string: 
#> [1] "bar"
fmt_txt("foo ", bold = TRUE) + fmt_txt("bar")
#> fmt_txt string: 
#> [1] "foo bar"
as.character(fmt_txt(2))
#> [1] "2"
```
