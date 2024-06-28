test_that("read_xml", {

  xml <- '<?xml test="yay" ?><a>A & B</a>'

  # TODO add test for isfile
  # and do some actual tests

  expect_silent(z <- readXMLPtr(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))
  expect_silent(z <- readXMLPtr(xml, isfile = FALSE, escapes = FALSE, declaration = TRUE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))
  expect_silent(z <- readXMLPtr(xml, isfile = FALSE, escapes = TRUE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))
  expect_silent(z <- readXMLPtr(xml, isfile = FALSE, escapes = TRUE, declaration = TRUE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))

  exp <- "<a>A & B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))

  exp <- "<?xml test=\"yay\"?><a>A & B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = FALSE, declaration = TRUE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))

  exp <- "<a>A &amp; B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = TRUE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))

  exp <- "<?xml test=\"yay\"?><a>A &amp; B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = TRUE, declaration = TRUE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))

  xml <- "<a>"
  expect_error(readXMLPtr(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))
  expect_error(readXML(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE))

})


test_that("xml_node", {

  xml_str <- "<a><b><c><d><e/></d></c></b></a>"
  xml <- read_xml(xml_str)

  exp <- xml_str
  expect_equal("a", getXMLXPtrName1(xml))
  expect_equal(exp, getXMLXPtr0(xml))
  expect_equal(exp, getXMLXPtr1(xml, "a"))

  exp <- "<b><c><d><e/></d></c></b>"
  expect_equal(exp, getXMLXPtr2(xml, "a", "b"))

  exp <- "<c><d><e/></d></c>"
  expect_equal(exp, getXMLXPtr3(xml, "a", "b", "c"))
  # bit cheating, this test returns the same, but not the actual feature of "*"
  expect_equal(exp, unkgetXMLXPtr3(xml, "a", "c"))

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

})

test_that("xml_attr", {

  exp <- list(c(a = "1"))

  xml_str <- "<a a=\"1\"/>"
  xml <- read_xml(xml_str)
  expect_equal(getXMLXPtr1attr(xml, "a"), exp)

  xml_str <- "<b><a a=\"1\"/></b>"
  xml <- read_xml(xml_str)
  expect_equal(getXMLXPtr2attr(xml, "b", "a"), exp)

  xml_str <- "<c><b><a a=\"1\"/></b></c>"
  xml <- read_xml(xml_str)
  expect_equal(getXMLXPtr3attr(xml, "c", "b", "a"), exp)

})

test_that("xml_append_child", {

  xml_node <- read_xml("<node><child1/><child2/></node>")
  xml_child <- read_xml("<new_child>&</new_child>")
  exp <- "<node><child1/><child2/><new_child>&</new_child></node>"
  expect_equal(xml_append_child1(xml_node, xml_child, pointer = FALSE), exp)

  # xml_node sets the flags for both
  xml_node <- read_xml("<node><child1/><child2/></node>", escapes = TRUE)
  xml_child <- read_xml("<new_child>&</new_child>")
  exp <- "<node><child1/><child2/><new_child>&amp;</new_child></node>"
  expect_equal(xml_append_child1(xml_node, xml_child, pointer = FALSE), exp)


  xml_node <- "<a><b/></a>"
  xml_child <- read_xml("<c/>")

  xml_node <- xml_append_child1(read_xml(xml_node), xml_child, pointer = FALSE)
  expect_equal(xml_node, "<a><b/><c/></a>")

  xml_node <- xml_append_child2(read_xml(xml_node), xml_child, level1 = "b", pointer = FALSE)
  expect_equal(xml_node, "<a><b><c/></b><c/></a>")

  xml_node <- xml_append_child3(read_xml(xml_node), read_xml("<d/>"), level1 = "b", level2 = "c", pointer = FALSE)
  expect_equal(xml_node, "<a><b><c><d/></c></b><c/></a>")


  # check that escapes does not throw a warning
  xml_node <- "<a><b/></a>"
  xml_child <- read_xml("<c>a&b</c>", escapes = FALSE)

  xml_node <- xml_append_child1(read_xml(xml_node, escapes = TRUE), xml_child, pointer = FALSE)
  expect_equal(xml_node, "<a><b/><c>a&amp;b</c></a>")

  xml_node <- xml_append_child2(read_xml(xml_node, escapes = TRUE), xml_child, level1 = "b", pointer = FALSE)
  expect_equal(xml_node, "<a><b><c>a&amp;b</c></b><c>a&amp;b</c></a>")

  xml_node <- xml_append_child3(read_xml(xml_node, escapes = TRUE), read_xml("<d/>"), level1 = "b", level2 = "c", pointer = FALSE)
  expect_equal(xml_node, "<a><b><c>a&amp;b<d/></c></b><c>a&amp;b</c></a>")

  # check that pointer is valid
  expect_s3_class(xml_append_child1(read_xml(xml_node), xml_child, pointer = TRUE), "pugi_xml")
  expect_s3_class(xml_append_child1(read_xml(xml_node), xml_child, pointer = TRUE), "pugi_xml")

  expect_s3_class(xml_append_child2(read_xml(xml_node), xml_child, level1 = "a", pointer = TRUE), "pugi_xml")
  expect_s3_class(xml_append_child2(read_xml(xml_node), xml_child, level1 = "a", pointer = TRUE), "pugi_xml")

  expect_s3_class(xml_append_child3(read_xml(xml_node), xml_child, level1 = "a", level2 = "b", pointer = TRUE), "pugi_xml")
  expect_s3_class(xml_append_child3(read_xml(xml_node), xml_child, level1 = "a", level2 = "b", pointer = TRUE), "pugi_xml")

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

test_that("is_xml", {

  expect_true(is_xml("<a/>"))
  expect_false(is_xml("a"))

})

test_that("getXMLPtr1con", {

  xml <- "<xml><a/><a/><b/></xml>"
  got <- getXMLPtr1con(read_xml(xml))
  exp <- c("<a/>", "<a/>", "<b/>")
  expect_equal(got, exp)

})

test_that("write_xmlPtr", {

  xml <- read_xml("<a/>")
  temp <- tempfile(fileext = ".xml")
  expect_silent(write_xmlPtr(xml, temp))

  skip_on_cran()
  expect_error(write_xmlPtr(xml, paste0("/", random_string(), "/test.xml")), "could not save file")
})
