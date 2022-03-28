test_that("read_xml", {

  xml <- '<?xml test="yay" ?><a>A & B</a>'

  # TODO add test for isfile
  # and do some actual tests

  expect_silent(z <- readXMLPtr(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE))
  expect_silent(z <- readXMLPtr(xml, isfile = FALSE, escapes = FALSE, declaration = TRUE))
  expect_silent(z <- readXMLPtr(xml, isfile = FALSE, escapes = TRUE, declaration = FALSE))
  expect_silent(z <- readXMLPtr(xml, isfile = FALSE, escapes = TRUE, declaration = TRUE))

  exp <- "<a>A & B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE))

  exp <- "<?xml test=\"yay\"?><a>A & B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = FALSE, declaration = TRUE))

  exp <- "<a>A &amp; B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = TRUE, declaration = FALSE))

  exp <- "<?xml test=\"yay\"?><a>A &amp; B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = TRUE, declaration = TRUE))

  xml <- "<a>"
  expect_error(readXMLPtr(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE))
  expect_error(readXML(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE))

})


test_that("xml_node", {

  xml_str <- "<a><b><c><d><e/></d></c></b></a>"
  xml <- read_xml(xml_str)

  exp <- xml_str
  expect_equal("a", getXMLXPtrName(xml))
  expect_equal(exp, getXMLXPtr0(xml))
  expect_equal(exp, getXMLXPtr1(xml, "a"))

  exp <- "<b><c><d><e/></d></c></b>"
  expect_equal(exp, getXMLXPtr2(xml, "a", "b"))

  exp <- "<c><d><e/></d></c>"
  expect_equal(exp, getXMLXPtr3(xml, "a", "b", "c"))
  # bit cheating, this test returns the same, but not the actual feature of "*"
  expect_equal(exp, unkgetXMLXPtr3(xml, "a", "c"))

  exp <- list(c("<d><e/></d>"))
  expect_equal(exp, getXMLXPtr4(xml, "a", "b", "c", "d"))

})

test_that("xml_value", {

  exp <- "1"

  xml_str <- "<a>1</a>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtr1val(xml, "a"))

  xml_str <- "<a><b>1</b></a>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtr2val(xml, "a", "b"))

  xml_str <- "<a><b><c>1</c></b></a>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtr3val(xml, "a", "b", "c"))

  xml_str <- "<a><b><c><d>1</d></c></b></a>"
  xml <- read_xml(xml_str)
  expect_equal(list(exp), getXMLXPtr4val(xml, "a", "b", "c", "d"))

})

test_that("xml_attr", {

  exp <- list(c(a="1"))

  xml_str <- "<a a=\"1\"/>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtr1attr(xml, "a"))

  xml_str <- "<b><a a=\"1\"/></b>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtr2attr(xml, "b", "a"))

  xml_str <- "<c><b><a a=\"1\"/></b></c>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtr3attr(xml, "c", "b", "a"))

  xml_str <- "<d><c><b><a a=\"1\"/></b></c></d>"
  xml <- read_xml(xml_str)
  expect_equal(list(list(as.list(unlist(exp)))), getXMLXPtr4attr(xml, "d", "c", "b", "a"))

})

test_that("xml_append_child", {

  xml_node <- read_xml("<node><child1/><child2/></node>")
  xml_child <- read_xml("<new_child>&</new_child>")
  exp <- "<node><child1/><child2/><new_child>&</new_child></node>"
  expect_equal(exp, xml_append_child(xml_node, xml_child, pointer = FALSE, escapes = FALSE))

  xml_node <- read_xml("<node><child1/><child2/></node>")
  xml_child <- read_xml("<new_child>&</new_child>")
  exp <- "<node><child1/><child2/><new_child>&amp;</new_child></node>"
  expect_equal(exp, xml_append_child(xml_node, xml_child, pointer = FALSE, escapes = TRUE))

  expect_true(inherits(xml_append_child(xml_node, xml_child, pointer = TRUE, escapes = FALSE), "pugi_xml"))
  expect_true(inherits(xml_append_child(xml_node, xml_child, pointer = TRUE, escapes = TRUE), "pugi_xml"))

})


test_that("set_sst", {

  exp <- "<si><t>a</t></si>"
  expect_equal(exp, set_sst("a"))

})

test_that("col2int and int2col", {

  test <- "B"
  x <- col2int(test)
  that <- int2col(x)

  expect_equal(test, that)

  test <- "AABWAWD"
  x <- col2int(test)
  that <- int2col(x)

  expect_equal(test, that)

  x <- col_to_int(test)
  that <- int_to_col(x)

  expect_equal(test, that)

})
