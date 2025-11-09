# read xml file

read xml file

## Usage

``` r
read_xml(
  xml,
  pointer = TRUE,
  escapes = FALSE,
  declaration = FALSE,
  whitespace = TRUE,
  empty_tags = FALSE,
  skip_control = TRUE
)
```

## Arguments

- xml:

  something to read character string or file

- pointer:

  should a pointer be returned?

- escapes:

  bool if characters like "&" should be escaped. The default is no
  escapes. Assuming that the input already provides valid information.

- declaration:

  should the declaration be imported

- whitespace:

  should whitespace pcdata be imported

- empty_tags:

  should `<b/>` or `<b></b>` be returned

- skip_control:

  should whitespace character be exported

## Details

Read xml files or strings to pointer and checks if the input is valid
XML. If the input is read into a character object, it will be
reevaluated every time it is called. A pointer is evaluated once, but
lives only for the lifetime of the R session or once it is gc().

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
  str(x)
#> Class 'pugi_xml' <externalptr> 
#>  - attr(*, "escapes")= logi FALSE
#>  - attr(*, "empty_tags")= logi FALSE
#>  - attr(*, "skip_control")= logi TRUE

  # a character
  y <- read_xml("<a><b/></a>", pointer = FALSE)
  print(y)
#> [1] "<a><b/></a>"
  print(y, raw = TRUE)
#> [1] "<a><b/></a>"
  str(y)
#>  chr "<a><b/></a>"

  # Errors if the import was unsuccessful
  try(z <- read_xml("<a><b/>"))
#> Error : xml import unsuccessful

  xml <- '<?xml test="yay" ?><a>A & B</a>'
  # difference in escapes
  read_xml(xml, escapes = TRUE, pointer = FALSE)
#> [1] "<a>A &amp; B</a>"
  read_xml(xml, escapes = FALSE, pointer = FALSE)
#> [1] "<a>A & B</a>"
  read_xml(xml, escapes = TRUE)
#> <a>A &amp; B</a>
  read_xml(xml, escapes = FALSE)
#> <a>A & B</a>

  # read declaration
  read_xml(xml, declaration = TRUE)
#> <?xml test="yay"?>
#> <a>A & B</a>
```
