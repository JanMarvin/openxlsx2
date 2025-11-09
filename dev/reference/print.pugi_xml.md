# print pugi_xml

print pugi_xml

## Usage

``` r
# S3 method for class 'pugi_xml'
print(x, indent = " ", raw = FALSE, attr_indent = FALSE, ...)
```

## Arguments

- x:

  something to print

- indent:

  indent used default is " "

- raw:

  print as raw text

- attr_indent:

  print attributes indented on new line

- ...:

  to please check

## Examples

``` r
  # a pointer
  x <- read_xml("<a><b/></a>")
  print(x)
#> <a>
#>  <b />
#> </a>
  print(x, raw = TRUE)
#> <a><b/></a>
```
