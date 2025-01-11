
test_that("read_xml", {

  exp <- "<a/>"
  got <- read_xml("<a/>", empty_tags = FALSE, pointer = FALSE)
  expect_equal(exp, got)

  exp <- "<a></a>"
  got <- read_xml("<a/>", empty_tags = TRUE, pointer = FALSE)
  expect_equal(exp, got)

  # a pointer
  x <- read_xml("<a><b/></a>")
  exp <- "<a>\n  <b/>\n</a>"

  expect_s3_class(x, "pugi_xml")

  xml <- "<a> </a>"
  got <- read_xml(xml, whitespace = TRUE, pointer = FALSE)
  expect_equal(xml, got)

  xml <- "<a> </a>"
  got <- read_xml(xml, whitespace = FALSE, pointer = FALSE)
  expect_equal("<a/>", got)

  xml <- "<a> <b> </b> </a>"
  got <- read_xml(xml, pointer = FALSE)
  expect_equal("<a><b> </b></a>", got)

  xml <- "<a> <b> </b> </a>"
  got <- paste(capture.output(read_xml(xml)), collapse = "\n")
  expect_equal("<a>\n <b> </b>\n</a>", got)

  xml <- "<a> <b> </b> </a>"
  got <- paste(capture.output(print(read_xml(xml), indent = "\t")), collapse = "\n")
  expect_equal("<a>\n\t<b> </b>\n</a>", got)

  # #does this even work?
  # expect_equal(cat(exp), print(x))

  # a character
  y <- read_xml("<a><b/></a>", pointer = FALSE)
  expect_type(y, "character")

  # Errors if the import was unsuccessful
  expect_error(z <- read_xml("<a><b/>"))
  # characters() are imported to <NA_character/> to avoid errors
  expect_equal("<NA_character_/>", read_xml(character(), pointer = FALSE))

  xml <- '<?xml test="yay"?><a>A & B</a>'
  # difference in escapes
  exp <- "<a>A &amp; B</a>"
  expect_equal(exp, read_xml(xml, escapes = TRUE, pointer = FALSE))
  exp <- "<a>A & B</a>"
  expect_equal(exp, read_xml(xml, escapes = FALSE, pointer = FALSE))

  # read declaration
  expect_equal(xml, read_xml(xml, declaration = TRUE, pointer = FALSE))

  exp <- '<t xml:space="preserve"> </t>'
  expect_equal(exp, read_xml(exp, pointer = FALSE))

  tmp <- tempfile(fileext = ".xml")
  write_file(body = exp, fl = tmp)

  exp <- "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><t xml:space=\"preserve\"> </t>"
  got <- readLines(tmp, warn = FALSE)
  expect_equal(exp, got)

  xml <- '<a><b><c1/></b></a><a><b><c2/></b></a>'

  exp <- c("<a><b><c1/></b></a>", "<a><b><c2/></b></a>")
  got <- xml_node(xml, "a")
  expect_equal(exp, got)

  exp <- c("<b><c1/></b>", "<b><c2/></b>")
  got <- xml_node(xml, "a", "b")
  expect_equal(exp, got)

  exp <- "<c1/>"
  got <- xml_node(xml, "a", "b", "c1")
  expect_equal(exp, got)

  exp <- "<c2/>"
  got <- xml_node(xml, "a", "b", "c2")
  expect_equal(exp, got)

})

test_that("xml_node", {

  xml <- "<a><b/></a>"
  x <- read_xml(xml, pointer = FALSE)

  expect_equal(xml_node_name(x), "a")
  expect_equal(xml_node(x), xml)
  expect_equal(xml_node(x, "a"), xml)
  expect_error(xml_node(x, 1))

  expect_equal(xml_node(x, "a", "b"), "<b/>")


  expect_equal(xml_node("<a><b/></a>", "a"), xml)

  expect_equal(xml_node("<a><b/></a>", "a", "b"), "<b/>")


  xml_str <- "<a><b><c><d><e/></d></c></b></a>"
  xml <- read_xml(xml_str)

  expect_equal(xml_node_name(xml_str, "a"), "b")
  expect_equal(xml_node_name(xml_str, "a", "b"), "c")
  expect_equal(xml_node_name(xml, "a"), "b")
  expect_equal(xml_node_name(xml, "a", "b"), "c")

  exp <- xml_str
  expect_equal(xml_node(xml, "a"), exp)

  exp <- "<b><c><d><e/></d></c></b>"
  expect_equal(xml_node(xml, "a", "b"), exp)

  exp <- "<c><d><e/></d></c>"
  expect_equal(xml_node(xml, "a", "b", "c"), exp)
  # bit cheating, this test returns the same, but not the actual feature of "*"
  expect_equal(xml_node(xml, "a", "*", "c"), exp)

})

test_that("xml_attr", {

  x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
  exp <- list(c(a = "1", b = "2"))
  expect_equal(xml_attr(x, "a"), exp)
  expect_error(xml_attr(x, 1))

  x <- read_xml("<a><b r=\"1\">2</b></a>")
  exp <- list(c(r = "1"))
  expect_equal(xml_attr(x, "a", "b"), exp)

  x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
  exp <- list(c(a = "1", b = "2"))
  expect_equal(xml_attr(x, "a"), exp)

  x <- read_xml("<b><a a=\"1\" b=\"2\"/></b>")
  exp <- list(c(a = "1", b = "2"))
  expect_equal(xml_attr(x, "b", "a"), exp)



  exp <- list(c(a = "1", b = "2"))
  expect_equal(xml_attr("<a a=\"1\" b=\"2\">1</a>", "a"), exp)

  exp <- list(c(r = "1"))
  expect_equal(xml_attr("<a><b r=\"1\">2</b></a>", "a", "b"), exp)

  exp <- list(c(a = "1", b = "2"))
  expect_equal(xml_attr("<a a=\"1\" b=\"2\">1</a>", "a"), exp)

  exp <- list(c(a = "1", b = "2"))
  expect_equal(xml_attr("<b><a a=\"1\" b=\"2\"/></b>", "b", "a"), exp)



  exp <- list(c(a = "1"))

  xml_str <- "<a a=\"1\"/>"
  xml <- read_xml(xml_str)
  expect_equal(xml_attr(xml, "a"), exp)

  xml_str <- "<b><a a=\"1\"/></b>"
  xml <- read_xml(xml_str)
  expect_equal(xml_attr(xml, "b", "a"), exp)

  xml_str <- "<c><b><a a=\"1\"/></b></c>"
  xml <- read_xml(xml_str)
  expect_equal(xml_attr(xml, "c", "b", "a"), exp)

})

test_that("xml_value", {

  x <- read_xml("<a>1</a>")
  expect_equal(xml_value(x, "a"), "1")
  expect_error(xml_value(x, 1))

  x <- read_xml("<a><b r=\"1\">2</b></a>")
  expect_equal(xml_value(x, "a", "b"), "2")

  x <- read_xml("<a><b r=\"1\">2</b><b r=\"2\">3</b></a>")
  expect_equal(xml_value(x, "a", "b"), c("2", "3"))


  exp <- "1"

  xml_str <- "<a>1</a>"
  xml <- read_xml(xml_str)
  expect_equal(xml_value(xml, "a"), "1")

  xml_str <- "<a><b>1</b></a>"
  xml <- read_xml(xml_str)
  expect_equal(xml_value(xml, "a", "b"), "1")

  xml_str <- "<a><b><c>1</c></b></a>"
  xml <- read_xml(xml_str)
  expect_equal(xml_value(xml, "a", "b", "c"), "1")

})

test_that("as_xml", {

  xml_str <- "<a><b><c><d>1</d></c></b></a>"

  # not the best test
  expect_equal(class(as_xml(xml_str)), "pugi_xml")

})

test_that("print", {

  xml_str <- "<a/>"

  expect_output(print(as_xml(xml_str)), "<a />")
  expect_output(print(as_xml(xml_str), raw = TRUE), "<a/>")

  xml_str <- '<a b1="foo" b2 = "bar"/>'
  expect_output(print(as_xml(xml_str), attr_indent = TRUE), '<a\n b1="foo"\n b2="bar" />')

})

test_that("xml_add_child", {

  xml_node <- "<node><child1/><child2/></node>"
  xml_child <- "<new_child/>"

  exp <- "<node><child1/><child2/><new_child/></node>"

  expect_equal(xml_add_child(xml_node, xml_child), exp)

  expect_error(xml_add_child(xml_node))
  expect_error(xml_add_child(xml_child = xml_child))

  xml_node <- "<a><b/></a>"
  xml_child <- "<c/>"

  xml_node <- xml_add_child(xml_node, xml_child)
  expect_equal(xml_node, "<a><b/><c/></a>")

  xml_node <- xml_add_child(xml_node, xml_child, level = c("b"))
  expect_equal(xml_node, "<a><b><c/></b><c/></a>")

  xml_node <- xml_add_child(xml_node, "<d/>", level = c("b", "c"))
  expect_equal(xml_node, "<a><b><c><d/></c></b><c/></a>")

})


test_that("xml_rm_child", {


  rm_child <- function(which) {
    xml_rm_child(
      xml_node = "<a><c>1</c><c>2</c></a>",
      xml_child = "c",
      which = which
    )
  }
  expect_equal(rm_child(which = 0), "<a/>")
  expect_equal(rm_child(which = 1), "<a><c>2</c></a>")
  expect_equal(rm_child(which = 2), "<a><c>1</c></a>")
  expect_equal(rm_child(which = 3), "<a><c>1</c><c>2</c></a>")

  xml_node <- "<a><b>1</b><b><c><d/></c><c/></b><c>2</c><c/></a>"
  xml_child <- "c"

  got <- xml_rm_child(xml_node, "b", which = 1)
  exp <- "<a><b><c><d/></c><c/></b><c>2</c><c/></a>"
  expect_equal(got, exp)

  xml_node <- exp

  got <- xml_rm_child(xml_node, xml_child, "b", which = 1)
  exp <- "<a><b><c/></b><c>2</c><c/></a>"
  expect_equal(got, exp)

  got <- xml_rm_child(xml_node, xml_child, level = "b", which = 2)
  exp <- "<a><b><c><d/></c></b><c>2</c><c/></a>"
  expect_equal(got, exp)

  got <- xml_rm_child(xml_node, xml_child, "b", which = 0)
  exp <- "<a><b/><c>2</c><c/></a>"
  expect_equal(got, exp)

  xml_node <- "<x><a><b><c>1</c><c>2</c><c>3</c></b></a></x>"

  got <- xml_rm_child(xml_node, xml_child, level = c("a", "b"), which = 2)
  exp <- "<x><a><b><c>1</c><c>3</c></b></a></x>"
  expect_equal(got, exp)

  got <- xml_rm_child(xml_node, xml_child, level = c("a", "b"), which = 0)
  exp <- "<x><a><b/></a></x>"
  expect_equal(got, exp)

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

  # only add node
  xml_node <- "<a foo=\"bar\">openxlsx2</a><b />"
  xml_attr <- c(foo = "", qux = "quux")
  xml_exp <- "<a foo=\"bar\" qux=\"quux\">openxlsx2</a><b qux=\"quux\"/>"
  xml_got <- xml_attr_mod(xml_node, xml_attr, remove_empty_attr = FALSE)
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

  xml_exp <- "<a><b/><c/></a>"
  xml_got <- xml_node_create("a", xml_children = c("<b/><c/>"))
  expect_identical(xml_exp, xml_got)

})

test_that("works with x namespace", {

  # create artificial xml file that will trigger x namespace removal
  tmp <- tempfile(fileext = ".xml")
  xml <- '<?xml xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ?><x:a><x:b/></x:a>'
  writeLines(xml, tmp)

  exp <- "<x:a><x:b/></x:a>"
  got <- read_xml(tmp, pointer = FALSE)
  expect_equal(exp, got)

  op <- options("openxlsx2.namespace_xml" = "x")
  on.exit(options(op), add = TRUE)

  exp <- "<a><b/></a>"
  got <- read_xml(tmp, pointer = FALSE)
  expect_equal(exp, got)

})
