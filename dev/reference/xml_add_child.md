# append xml child to node

append xml child to node

## Usage

``` r
xml_add_child(xml_node, xml_child, level, pointer = FALSE, ...)
```

## Arguments

- xml_node:

  xml_node

- xml_child:

  xml_child

- level:

  optional level, if missing the first child is picked

- pointer:

  pointer

- ...:

  additional arguments passed to
  [`read_xml()`](https://janmarvin.github.io/openxlsx2/dev/reference/read_xml.md)

## Examples

``` r
xml_node <- "<a><b/></a>"
xml_child <- "<c/>"

# add child to first level node
xml_add_child(xml_node, xml_child)
#> [1] "<a><b/><c/></a>"

# add child to second level node as request
xml_node <- xml_add_child(xml_node, xml_child, level = c("b"))

# add child to third level node as request
xml_node <- xml_add_child(xml_node, "<d/>", level = c("b", "c"))
```
