
test_that("xml_attr_mod", {

  # add single node
  xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
  xml_attr <- c(qux = "quux")
  xml_exp <- "<a foo=\"bar\" qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
  xml_got <- xml_attr_mod(xml_node, xml_attr)
  expect_identical(xml_exp, xml_got)

  # update node and add node
  xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
  xml_attr <- c(foo = "baz", qux = "quux")
  xml_exp <- "<a foo=\"baz\" qux=\"quux\">openxlsx2</a><b foo=\"baz\" qux=\"quux\"/>"
  xml_got <- xml_attr_mod(xml_node, xml_attr)
  expect_identical(xml_exp, xml_got)

  # remove node and add node
  xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
  xml_attr <- c(foo = "", qux = "quux")
  xml_exp <- "<a qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
  xml_got <- xml_attr_mod(xml_node, xml_attr)
  expect_identical(xml_exp, xml_got)

})

test_that("xml_node_create", {

  # create node
  xml_name <- "a"
  xml_exp <- "<a/>"
  xml_got <- xml_node_create(xml_name)
  expect_identical(xml_exp, xml_got)

  # add child
  xml_child <- "openxlsx"
  xml_exp <- "<a>openxlsx</a>"
  xml_got <- xml_node_create(xml_name, xml_children = xml_child)
  expect_identical(xml_exp, xml_got)

  # add attributes
  xml_attr <- c(foo = "baz", qux = "quux")
  xml_exp <- "<a foo=\"baz\" qux=\"quux\"/>"
  xml_got <- xml_node_create(xml_name, xml_attributes = xml_attr)
  expect_identical(xml_exp, xml_got)

  # add child and attributes
  xml_exp <- "<a foo=\"baz\" qux=\"quux\">openxlsx</a>"
  xml_got <- xml_node_create(xml_name, xml_children = xml_child, xml_attributes = xml_attr)
  expect_identical(xml_exp, xml_got)

})
