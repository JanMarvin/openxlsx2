# xml_node

returns xml values as character

## Usage

``` r
xml_node(xml, level1 = NULL, level2 = NULL, level3 = NULL, ...)

xml_node_name(xml, level1 = NULL, level2 = NULL, ...)

xml_value(xml, level1 = NULL, level2 = NULL, level3 = NULL, ...)

xml_attr(xml, level1 = NULL, level2 = NULL, level3 = NULL, ...)
```

## Arguments

- xml:

  something xml

- level1:

  to please check

- level2:

  to please check

- level3:

  to please check

- ...:

  additional arguments passed to
  [`read_xml()`](https://janmarvin.github.io/openxlsx2/reference/read_xml.md)

## Details

This function returns XML nodes as used in openxlsx2. In theory they
could be returned as pointers as well, but this has not yet been
implemented. If no level is provided, the nodes on level1 are returned

## Examples

``` r
  x <- read_xml("<a><b/></a>")
  # return a
  xml_node(x, "a")
#> [1] "<a><b/></a>"
  # return b. requires the path to the node
  xml_node(x, "a", "b")
#> [1] "<b/>"
  xml_node_name("<a/>")
#> [1] "a"
  xml_node_name("<a><b/></a>", "a")
#> [1] "b"
  x <- read_xml("<a>1</a>")
  xml_value(x, "a")
#> [1] "1"

  x <- read_xml("<a><b r=\"1\">2</b></a>")
  xml_value(x, "a", "b")
#> [1] "2"

  x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
  xml_attr(x, "a")
#> [[1]]
#>   a   b 
#> "1" "2" 
#> 

  x <- read_xml("<a><b r=\"1\">2</b></a>")
  xml_attr(x, "a", "b")
#> [[1]]
#>   r 
#> "1" 
#> 
  x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
  xml_attr(x, "a")
#> [[1]]
#>   a   b 
#> "1" "2" 
#> 

  x <- read_xml("<b><a a=\"1\" b=\"2\"/></b>")
  xml_attr(x, "b", "a")
#> [[1]]
#>   a   b 
#> "1" "2" 
#> 
```
