test_that("read_xml", {

  xml <- '<?xml test="yay" ?><a>A & B</a>'

  # TODO add test for isfile
  # and do some actual tests

  expect_silent(z <- readXML(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = TRUE))
  expect_silent(z <- readXML(xml, isfile = FALSE, escapes = FALSE, declaration = TRUE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = TRUE))
  expect_silent(z <- readXML(xml, isfile = FALSE, escapes = TRUE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = TRUE))
  expect_silent(z <- readXML(xml, isfile = FALSE, escapes = TRUE, declaration = TRUE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = TRUE))

  exp <- "<a>A & B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = FALSE))

  exp <- "<?xml test=\"yay\"?><a>A & B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = FALSE, declaration = TRUE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = FALSE))

  exp <- "<a>A &amp; B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = TRUE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = FALSE))

  exp <- "<?xml test=\"yay\"?><a>A &amp; B</a>"
  expect_equal(exp, readXML(xml, isfile = FALSE, escapes = TRUE, declaration = TRUE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = FALSE))

  xml <- "<a>"
  expect_error(readXML(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = TRUE))
  expect_error(readXML(xml, isfile = FALSE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE, pointer = FALSE))

})


test_that("xml_node", {

  xml_str <- "<a><b><c><d><e/></d></c></b></a>"
  xml <- read_xml(xml_str)

  exp <- xml_str
  expect_equal("a", getXMLXPtrNamePath(xml, character()))
  expect_equal(exp, getXMLXPtrPath(xml, character()))
  expect_equal(exp, getXMLXPtrPath(xml, "a"))

  exp <- "<b><c><d><e/></d></c></b>"
  expect_equal(exp, getXMLXPtrPath(xml, c("a", "b")))

  exp <- "<c><d><e/></d></c>"
  expect_equal(exp, getXMLXPtrPath(xml, c("a", "b", "c")))

  # bit cheating, this test returns the same, but not the actual feature of "*"
  expect_equal(exp, getXMLXPtrPath(xml, c("a", "*", "c")))

})

test_that("xml_value", {

  exp <- "1"

  xml_str <- "<a>1</a>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtrValPath(xml, "a"))

  xml_str <- "<a><b>1</b></a>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtrValPath(xml, c("a", "b")))

  xml_str <- "<a><b><c>1</c></b></a>"
  xml <- read_xml(xml_str)
  expect_equal(exp, getXMLXPtrValPath(xml, c("a", "b", "c")))

})

test_that("xml_attr", {

  exp <- list(c(a = "1"))

  xml_str <- "<a a=\"1\"/>"
  xml <- read_xml(xml_str)
  expect_equal(getXMLXPtrAttrPath(xml, "a"), exp)

  xml_str <- "<b><a a=\"1\"/></b>"
  xml <- read_xml(xml_str)
  expect_equal(getXMLXPtrAttrPath(xml, c("b", "a")), exp)

  xml_str <- "<c><b><a a=\"1\"/></b></c>"
  xml <- read_xml(xml_str)
  expect_equal(getXMLXPtrAttrPath(xml, c("c", "b", "a")), exp)

})

test_that("xml_append_child", {

  xml_node <- read_xml("<node><child1/><child2/></node>")
  xml_child <- read_xml("<new_child>&</new_child>")
  exp <- "<node><child1/><child2/><new_child>&</new_child></node>"
  level <- character()
  expect_equal(xml_append_child_path(xml_node, xml_child, level, pointer = FALSE), exp)

  # xml_node sets the flags for both
  xml_node <- read_xml("<node><child1/><child2/></node>", escapes = TRUE)
  xml_child <- read_xml("<new_child>&</new_child>")
  exp <- "<node><child1/><child2/><new_child>&amp;</new_child></node>"
  level <- character()
  expect_equal(xml_append_child_path(xml_node, xml_child, level, pointer = FALSE), exp)


  xml_node <- "<a><b/></a>"
  xml_child <- read_xml("<c/>")
  level <- character()

  xml_node <- xml_append_child_path(read_xml(xml_node), xml_child, level, pointer = FALSE)
  expect_equal(xml_node, "<a><b/><c/></a>")

  xml_node <- xml_append_child_path(read_xml(xml_node), xml_child, c(level1 = "b"), pointer = FALSE)
  expect_equal(xml_node, "<a><b><c/></b><c/></a>")

  xml_node <- xml_append_child_path(read_xml(xml_node), read_xml("<d/>"), c(level1 = "b", level2 = "c"), pointer = FALSE)
  expect_equal(xml_node, "<a><b><c><d/></c></b><c/></a>")


  # check that escapes does not throw a warning
  xml_node <- "<a><b/></a>"
  xml_child <- read_xml("<c>a&b</c>", escapes = FALSE)
  level <- character()

  xml_node <- xml_append_child_path(read_xml(xml_node, escapes = TRUE), xml_child, level, pointer = FALSE)
  expect_equal(xml_node, "<a><b/><c>a&amp;b</c></a>")

  xml_node <- xml_append_child_path(read_xml(xml_node, escapes = TRUE), xml_child, c(level1 = "b"), pointer = FALSE)
  expect_equal(xml_node, "<a><b><c>a&amp;b</c></b><c>a&amp;b</c></a>")

  xml_node <- xml_append_child_path(read_xml(xml_node, escapes = TRUE), read_xml("<d/>"), c(level1 = "b", level2 = "c"), pointer = FALSE)
  expect_equal(xml_node, "<a><b><c>a&amp;b<d/></c></b><c>a&amp;b</c></a>")

  # check that pointer is valid
  expect_s3_class(xml_append_child_path(read_xml(xml_node), xml_child, level, pointer = TRUE), "pugi_xml")
  expect_s3_class(xml_append_child_path(read_xml(xml_node), xml_child, level, pointer = TRUE), "pugi_xml")

  expect_s3_class(xml_append_child_path(read_xml(xml_node), xml_child, c(level1 = "a"), pointer = TRUE), "pugi_xml")
  expect_s3_class(xml_append_child_path(read_xml(xml_node), xml_child, c(level1 = "a"), pointer = TRUE), "pugi_xml")

  expect_s3_class(xml_append_child_path(read_xml(xml_node), xml_child, c(level1 = "a", level2 = "b"), pointer = TRUE), "pugi_xml")
  expect_s3_class(xml_append_child_path(read_xml(xml_node), xml_child, c(level1 = "a", level2 = "b"), pointer = TRUE), "pugi_xml")

})

test_that("col2int and int2col", {

  test <- "B"
  x <- col2int(test)
  that <- int2col(x)

  expect_equal(test, that)

  test <- "AABWAWD"
  expect_error(col2int(test), "Column exceeds valid range")
  expect_error(int2col(322116630L), "Column exceeds valid range")

  expect_error(col_to_int(test), "Column exceeds valid range")
  expect_error(ox_int_to_col(322116630L), "Column exceeds valid range")

  ## fuzzing input
  expect_equal(col2int(integer()), integer())
  expect_equal(col2int(character()), integer())
  expect_equal(col2int(NULL), NULL)
  expect_error(col2int(""), "Empty string in conversion from column to integer")
  expect_error(col2int("_"), "found non alphabetic character in column to integer conversion")
  expect_error(col2int(" "), "found non alphabetic character in column to integer conversion")
  expect_error(col2int("   "), "found non alphabetic character in column to integer conversion")
  expect_error(col2int(NA_character_), "x contains NA")
  expect_error(col2int(16385L), "Column exceeds valid range")
  expect_error(col2int(0L), "Column exceeds valid range")
  expect_error(col2int(-1L), "Column exceeds valid range")

  expect_error(int2col(-1), "Column exceeds valid range")
  expect_error(int2col(0), "Column exceeds valid range")
  expect_error(int2col(16385), "Column exceeds valid range")
  expect_equal(int2col(NULL), NULL)
  expect_equal(int2col(integer()), character())

})

test_that("is_xml", {

  expect_true(is_xml("<a/>"))
  expect_false(is_xml("a"))

})

test_that("getXMLXPtrContent", {

  xml <- "<xml><a/><a/><b/></xml>"
  got <- xml_node(xml, "*", "*")
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
