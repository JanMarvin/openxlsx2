# as.character pugi_xml

as.character pugi_xml

## Usage

``` r
# S3 method for class 'pugi_xml'
as.character(x, ...)
```

## Arguments

- x:

  a class pugi_xml object

- ...:

  additional arguments see
  [`print.pugi_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/print.pugi_xml.md)

## Examples

``` r
  # a pointer
  x <- read_xml("<a><b/></a>")
  as.character(x)
#> [1] "<a><b/></a>"
  as.character(x, raw = FALSE)
#> [1] "<a>\n <b />\n</a>\n"
```
