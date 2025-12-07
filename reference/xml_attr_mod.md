# adds or updates attribute(s) in existing xml node

Needs xml node and named character vector as input. Modifies the
arguments of each first child found in the xml node and adds or updates
the attribute vector.

## Usage

``` r
xml_attr_mod(
  xml_content,
  xml_attributes,
  escapes = FALSE,
  declaration = FALSE,
  remove_empty_attr = TRUE
)
```

## Arguments

- xml_content:

  some valid xml_node

- xml_attributes:

  R vector of named attributes

- escapes:

  bool if escapes should be used

- declaration:

  bool if declaration should be imported

- remove_empty_attr:

  bool remove empty attributes or ignore them

## Details

If a named attribute in `xml_attributes` is "" remove the attribute from
the node. If `xml_attributes` contains a named entry found in the xml
node, it is updated else it is added as attribute.

## Examples

``` r
  # add single node
    xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
    xml_attr <- c(qux = "quux")
    # "<a foo=\"bar\" qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
    xml_attr_mod(xml_node, xml_attr)
#> [1] "<a foo=\"bar\" qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"

  # update node and add node
    xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
    xml_attr <- c(foo = "baz", qux = "quux")
    # "<a foo=\"baz\" qux=\"quux\">openxlsx2</a><b foo=\"baz\" qux=\"quux\"/>"
    xml_attr_mod(xml_node, xml_attr)
#> [1] "<a foo=\"baz\" qux=\"quux\">openxlsx2</a><b foo=\"baz\" qux=\"quux\"/>"

  # remove node and add node
    xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
    xml_attr <- c(foo = "", qux = "quux")
    # "<a qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
    xml_attr_mod(xml_node, xml_attr)
#> [1] "<a qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
```
