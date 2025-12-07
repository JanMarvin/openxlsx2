# remove xml child to node

remove xml child to node

## Usage

``` r
xml_rm_child(xml_node, xml_child, level, which = 0, pointer = FALSE, ...)
```

## Arguments

- xml_node:

  xml_node

- xml_child:

  xml_child

- level:

  optional level, if missing the first child is picked

- which:

  optional index which node to remove, if multiple are available.
  Default disabled all will be removed

- pointer:

  pointer

- ...:

  additional arguments passed to
  [`read_xml()`](https://janmarvin.github.io/openxlsx2/reference/read_xml.md)

## Examples

``` r
xml_node <- "<a><b><c><d/></c></b><c/></a>"
xml_child <- "c"

xml_rm_child(xml_node, xml_child)
#> [1] "<a><b><c><d/></c></b></a>"

xml_rm_child(xml_node, xml_child, level = c("b"))
#> [1] "<a><b/><c/></a>"

xml_rm_child(xml_node, "d", level = c("b", "c"))
#> [1] "<a><b><c/></b><c/></a>"
```
