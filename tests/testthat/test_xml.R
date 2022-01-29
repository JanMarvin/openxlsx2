
test_that("read_xml", {

  # a pointer
  x <- read_xml("<a><b/></a>")
  exp <- "<a>\n  <b/>\n</a>"

  expect_true(inherits(x, "pugi_xml"))

  # #does this even work?
  # expect_equal(cat(exp), print(x))

  # a character
  y <- read_xml("<a><b/></a>", pointer = FALSE)
  expect_true(is.character(y))

  # Errors if the import was unsuccessful
  expect_error(z <- read_xml("<a><b/>"))

  xml <- '<?xml test="yay"?><a>A & B</a>'
  # difference in escapes
  exp <- "<a>A &amp; B</a>"
  expect_equal(exp, read_xml(xml, escapes = TRUE, pointer = FALSE))
  exp <- "<a>A & B</a>"
  expect_equal(exp, read_xml(xml, escapes = FALSE, pointer = FALSE))

  # read declaration
  expect_equal(xml, read_xml(xml, declaration = TRUE, pointer = FALSE))

})

test_that("xml_node", {

  xml <- "<a><b/></a>"
  x <- read_xml(xml, pointer = FALSE)

  expect_equal(xml, xml_node(x, "a"))

  exp <- "<b/>"
  expect_equal(exp, xml_node(x, "a", "b"))


  expect_equal(xml, xml_node("<a><b/></a>", "a"))

  exp <- "<b/>"
  expect_equal(exp, xml_node("<a><b/></a>", "a", "b"))

})

test_that("xml_attribute", {

  x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
  exp <- list(c(a = "1", b = "2"))
  expect_equal(exp, xml_attr(x, "a"))

  x <- read_xml("<a><b r=\"1\">2</b></a>")
  exp <- list(c(r = "1"))
  expect_equal(exp, xml_attr(x, "a", "b"))

  x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
  exp <- list(c(a = "1", b = "2"))
  expect_equal(exp, xml_attr(x, "a"))

  x <- read_xml("<b><a a=\"1\" b=\"2\"/></b>")
  exp <- list(c(a = "1", b = "2"))
  expect_equal(exp, xml_attr(x, "b", "a"))



  exp <- list(c(a = "1", b = "2"))
  expect_equal(exp, xml_attr("<a a=\"1\" b=\"2\">1</a>", "a"))

  exp <- list(c(r = "1"))
  expect_equal(exp, xml_attr("<a><b r=\"1\">2</b></a>", "a", "b"))

  exp <- list(c(a = "1", b = "2"))
  expect_equal(exp, xml_attr("<a a=\"1\" b=\"2\">1</a>", "a"))

  exp <- list(c(a = "1", b = "2"))
  expect_equal(exp, xml_attr("<b><a a=\"1\" b=\"2\"/></b>", "b", "a"))

})

test_that("xml_value", {

  x <- read_xml("<a>1</a>")
  exp <- "1"
  expect_equal(exp, xml_value(x, "a"))

  x <- read_xml("<a><b r=\"1\">2</b></a>")
  exp <- "2"
  expect_equal(exp, xml_value(x, "a", "b"))

  x <- read_xml("<a><b r=\"1\">2</b><b r=\"2\">3</b></a>")
  exp <- c("2", "3")
  expect_equal(exp, xml_value(x, "a", "b"))

})

test_that("as_xml", {

  ## Not run:
  tmp_xlsx <- tempdir()
  xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
  unzip(xlsxFile, exdir = tmp_xlsx)

  wb <- loadWorkbook(xlsxFile)
  styles_xml <- sprintf("%s/xl/styles.xml", tmp_xlsx)

  # is external pointer
  sxml <- read_xml(styles_xml)

  # is character
  font <- xml_node(sxml, "styleSheet", "fonts", "font")

  ### no clue how to test this :-)
  # # is again external pointer
  # as_xml(font)

})

test_that("xml_attr_mod", {

  xml_node <- "<node><child1/><child2/></node>"
  xml_child <- "<new_child/>"

  exp <- "<node><child1/><child2/><new_child/></node>"

  expect_equal(exp, xml_add_child(xml_node, xml_child))

})


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
