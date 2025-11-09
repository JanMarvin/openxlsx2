# create xml_node from R objects

takes xml_name, xml_children and xml_attributes to create a new
xml_node.

## Usage

``` r
xml_node_create(
  xml_name,
  xml_children = NULL,
  xml_attributes = NULL,
  escapes = FALSE,
  declaration = FALSE
)
```

## Arguments

- xml_name:

  the name of the new xml_node

- xml_children:

  character vector children attached to the xml_node

- xml_attributes:

  named character vector of attributes for the xml_node

- escapes:

  bool if escapes should be used

- declaration:

  bool if declaration should be imported

## Details

if xml_children or xml_attributes should be empty, use NULL

## Examples

``` r
xml_name <- "a"
# "<a/>"
xml_node_create(xml_name)
#> [1] "<a/>"

xml_child <- "openxlsx"
# "<a>openxlsx</a>"
xml_node_create(xml_name, xml_children = xml_child)
#> [1] "<a>openxlsx</a>"

xml_attr <- c(foo = "baz", qux = "quux")
# "<a foo=\"baz\" qux=\"quux\"/>"
xml_node_create(xml_name, xml_attributes = xml_attr)
#> [1] "<a foo=\"baz\" qux=\"quux\"/>"

# "<a foo=\"baz\" qux=\"quux\">openxlsx</a>"
xml_node_create(xml_name, xml_children = xml_child, xml_attributes = xml_attr)
#> [1] "<a foo=\"baz\" qux=\"quux\">openxlsx</a>"
```
