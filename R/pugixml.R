#' read xml file
#' @name read_xml
#' @param xml something to read character string or file
#' @param declaration should the declaration be imported
#' @param escapes bool if characters like "&" should be escaped. The default is
#' no escapes. Assuming that the input already provides valid information.
#' @param whitespace should whitespace pcdata be imported
#' @param empty_tags should `<b/>` or `<b></b>` be returned
#' @param skip_control should whitespace character be exported
#' @details Read xml files or strings to pointer and checks if the input is
#' valid XML.
#' If the input is read into a character object, it will be reevaluated every
#' time it is called. A pointer is evaluated once, but lives only for the
#' lifetime of the R session or once it is gc().
#' @param pointer should a pointer be returned?
#' @examples
#'   # a pointer
#'   x <- read_xml("<a><b/></a>")
#'   print(x)
#'   print(x, raw = TRUE)
#'   str(x)
#'
#'   # a character
#'   y <- read_xml("<a><b/></a>", pointer = FALSE)
#'   print(y)
#'   print(y, raw = TRUE)
#'   str(y)
#'
#'   # Errors if the import was unsuccessful
#'   try(z <- read_xml("<a><b/>"))
#'
#'   xml <- '<?xml test="yay" ?><a>A & B</a>'
#'   # difference in escapes
#'   read_xml(xml, escapes = TRUE, pointer = FALSE)
#'   read_xml(xml, escapes = FALSE, pointer = FALSE)
#'   read_xml(xml, escapes = TRUE)
#'   read_xml(xml, escapes = FALSE)
#'
#'   # read declaration
#'   read_xml(xml, declaration = TRUE)
#'
#' @export
read_xml <- function(xml, pointer = TRUE, escapes = FALSE, declaration = FALSE, whitespace = TRUE, empty_tags = FALSE, skip_control = TRUE) {

  z <- NULL
  xml <- unclass(xml)

  isvml <- grepl("^.vml$", xml)

  isfile <- FALSE
  if (length(xml) == 1 && !to_long(xml) && file.exists(xml))
    isfile <- TRUE

  if (!isfile)
    xml <- stringi::stri_join(xml, collapse = "")

  if (identical(xml, ""))
    xml <- "<NA_character_ />"

  # Clean xml files from xml namespace. Otherwise all nodes might look like
  # <x:node/> and not <node/>. https://github.com/JanMarvin/openxlsx2/pull/213
  xml_ns <- getOption("openxlsx2.namespace_xml")
  if (!is.null(xml_ns) && isfile && !isvml) {

    xml_file <- stringi::stri_join(
      stringi::stri_read_lines(xml, encoding = "UTF-8"),
      collapse = "")

    if (grepl(sprintf('xmlns:%s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', xml_ns), xml_file, fixed = TRUE)) {
      xml_file <- stringi::stri_replace_all_fixed(xml_file, sprintf("<%s:", xml_ns), "<")
      xml_file <- stringi::stri_replace_all_fixed(xml_file, sprintf("</%s:", xml_ns), "</")

      # replace xml with already and cleaned output
      xml <- xml_file
      isfile <- FALSE
    }
  }

  if (pointer) {
    z <- readXMLPtr(xml, isfile, escapes, declaration, whitespace, empty_tags, skip_control)
  } else {
    z <- readXML(xml, isfile, escapes, declaration, whitespace, empty_tags, skip_control)
  }

  z
}


#' xml_node
#' @name pugixml
#' @param xml something xml
#' @param level1 to please check
#' @param level2 to please check
#' @param level3 to please check
#' @param ... additional arguments passed to `read_xml()`
#' @details This function returns XML nodes as used in openxlsx2. In theory they
#' could be returned as pointers as well, but this has not yet been implemented.
#' If no level is provided, the nodes on level1 are returned
#' @examples
#'   x <- read_xml("<a><b/></a>")
#'   # return a
#'   xml_node(x, "a")
#'   # return b. requires the path to the node
#'   xml_node(x, "a", "b")
#' @export
xml_node <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL, ...) {

  lvl <- c(level1, level2, level3)
  if (!all(is.null(lvl))) {
    lvl <- lvl[!is.null(lvl)]
    if (!all(is.character(lvl)))
      stop("levels must be character vectors")
  }

  z <- NULL

  if (!inherits(xml, "pugi_xml"))
    xml <- read_xml(xml, ...)


  if (inherits(xml, "pugi_xml")) {
    if (length(lvl) == 0) z <- getXMLXPtr0(xml)
    if (length(lvl) == 1) z <- getXMLXPtr1(xml, level1)
    if (length(lvl) == 2) z <- getXMLXPtr2(xml, level1, level2)
    if (length(lvl) == 3) z <- getXMLXPtr3(xml, level1, level2, level3)
    if (length(lvl) == 3) if (level2 == "*") z <- unkgetXMLXPtr3(xml, level1, level3)
  }

  z
}

#' @rdname pugixml
#' @examples
#'   xml_node_name("<a/>")
#'   xml_node_name("<a><b/></a>", "a")
#' @export
xml_node_name <- function(xml, level1 = NULL, level2 = NULL, ...) {
  lvl <- c(level1, level2)
  if (!inherits(xml, "pugi_xml")) xml <- read_xml(xml, ...)
  if (length(lvl) == 0) z <- getXMLXPtrName1(xml)
  if (length(lvl) == 1) z <- getXMLXPtrName2(xml, level1)
  if (length(lvl) == 2) z <- getXMLXPtrName3(xml, level1, level2)
  z
}

#' xml_value
#' @rdname pugixml
#' @description returns xml values as character
#' @examples
#'   x <- read_xml("<a>1</a>")
#'   xml_value(x, "a")
#'
#'   x <- read_xml("<a><b r=\"1\">2</b></a>")
#'   xml_value(x, "a", "b")
#' @export
xml_value <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL, ...) {

  lvl <- c(level1, level2, level3)
  lvl <- lvl[!is.null(lvl)]
  if (!all(is.character(lvl)))
    stop("levels must be character vectors")

  z <- NULL

  if (!inherits(xml, "pugi_xml"))
    xml <- read_xml(xml, ...)

  if (inherits(xml, "pugi_xml")) {
    if (length(lvl) == 1) z <- getXMLXPtr1val(xml, level1)
    if (length(lvl) == 2) z <- getXMLXPtr2val(xml, level1, level2)
    if (length(lvl) == 3) z <- getXMLXPtr3val(xml, level1, level2, level3)
  }

  z
}

#' @rdname pugixml
#' @examples
#'
#'   x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
#'   xml_attr(x, "a")
#'
#'   x <- read_xml("<a><b r=\"1\">2</b></a>")
#'   xml_attr(x, "a", "b")
#'   x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
#'   xml_attr(x, "a")
#'
#'   x <- read_xml("<b><a a=\"1\" b=\"2\"/></b>")
#'   xml_attr(x, "b", "a")
#' @export
xml_attr <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL,  ...) {

  lvl <- c(level1, level2, level3)
  lvl <- lvl[!is.null(lvl)]
  if (!all(is.character(lvl)))
    stop("levels must be character vectors")

  z <- NULL

  if (!inherits(xml, "pugi_xml"))
    xml <- read_xml(xml, ...)

  if (inherits(xml, "pugi_xml")) {
    if (length(lvl) == 1) z <- getXMLXPtr1attr(xml, level1)
    if (length(lvl) == 2) z <- getXMLXPtr2attr(xml, level1, level2)
    if (length(lvl) == 3) z <- getXMLXPtr3attr(xml, level1, level2, level3)
  }

  z
}

#' print pugi_xml
#' @method print pugi_xml
#' @param x something to print
#' @param indent indent used default is " "
#' @param raw print as raw text
#' @param attr_indent print attributes indented on new line
#' @param ... to please check
#' @examples
#'   # a pointer
#'   x <- read_xml("<a><b/></a>")
#'   print(x)
#'   print(x, raw = TRUE)
#' @export
print.pugi_xml <- function(x, indent = " ", raw = FALSE, attr_indent = FALSE, ...) {
  cat(printXPtr(x, indent, raw, attr_indent))
  if (raw) cat("\n")
}

#' loads character string to pugixml and returns an externalptr
#' @details
#' might be useful for larger documents where single nodes are shortened
#' and otherwise the full tree has to be reimported. unsure where we have
#' such a case.
#' is useful, for printing nodes from a larger tree, that have been exported
#' as characters (at some point in time we have to convert the xml to R)
#' @param x input as xml
#' @param ... additional arguments passed to `read_xml()`
#' @examples
#' tmp_xlsx <- tempfile()
#' xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' unzip(xlsxFile, exdir = tmp_xlsx)
#'
#' wb <- wb_load(xlsxFile)
#' styles_xml <- sprintf("%s/xl/styles.xml", tmp_xlsx)
#'
#' # is external pointer
#' sxml <- read_xml(styles_xml)
#'
#' # is character
#' font <- xml_node(sxml, "styleSheet", "fonts", "font")
#'
#' # is again external pointer
#' as_xml(font)
#' @export
as_xml <- function(x, ...) {
  read_xml(paste(x, collapse = ""), ...)
}

#' write xml file
#' @description brings the added benefit of xml checking
#' @param head head part of xml
#' @param body body part of xml
#' @param tail tail part of xml
#' @param fl file name with full path
#' @param escapes bool if characters like "&" should be escaped. The default is
#' no escape, assuming that xml to export is already valid.
#' @export
# TODO needs a unit test
write_file <- function(head = "", body = "", tail = "", fl = "", escapes = FALSE) {
  if (getOption("openxlsx2.force_utf8_encoding", default = FALSE)) {
      from_enc <- getOption("openxlsx2.native_encoding")
      head <- stringi::stri_encode(head, from = from_enc, to = "UTF-8")
      body <- stringi::stri_encode(body, from = from_enc, to = "UTF-8")
      tail <- stringi::stri_encode(tail, from = from_enc, to = "UTF-8")
  }
  xml_content <- stringi::stri_join(head, body, tail, collapse = "")
  xml_content <- write_xml_file(xml_content = xml_content, escapes = escapes)
  write_xmlPtr(xml_content, fl)
}

#' append xml child to node
#' @param xml_node xml_node
#' @param xml_child xml_child
#' @param level optional level, if missing the first child is picked
#' @param pointer pointer
#' @param ... additional arguments passed to `read_xml()`
#' @examples
#' xml_node <- "<a><b/></a>"
#' xml_child <- "<c/>"
#'
#' # add child to first level node
#' xml_add_child(xml_node, xml_child)
#'
#' # add child to second level node as request
#' xml_node <- xml_add_child(xml_node, xml_child, level = c("b"))
#'
#' # add child to third level node as request
#' xml_node <- xml_add_child(xml_node, "<d/>", level = c("b", "c"))
#'
#' @export
xml_add_child <- function(xml_node, xml_child, level, pointer = FALSE, ...) {

  if (missing(xml_node))
    stop("need xml_node")

  if (missing(xml_child))
    stop("need xml_child")

  if (all(xml_child == "")) return(xml_node)

  xml_node <- read_xml(xml_node, ...)
  xml_child <- read_xml(xml_child, ...)

  if (missing(level)) {
    z <- xml_append_child1(xml_node, xml_child, pointer)
  } else {

    if (length(level) == 1)
      z <- xml_append_child2(xml_node, xml_child, level[[1]], pointer)

    if (length(level) == 2)
      z <- xml_append_child3(xml_node, xml_child, level[[1]], level[[2]], pointer)

  }

  return(z)
}


#' remove xml child to node
#' @param xml_node xml_node
#' @param xml_child xml_child
#' @param level optional level, if missing the first child is picked
#' @param which optional index which node to remove, if multiple are available. Default disabled all will be removed
#' @param pointer pointer
#' @param ... additional arguments passed to `read_xml()`
#' @examples
#' xml_node <- "<a><b><c><d/></c></b><c/></a>"
#' xml_child <- "c"
#'
#' xml_rm_child(xml_node, xml_child)
#'
#' xml_rm_child(xml_node, xml_child, level = c("b"))
#'
#' xml_rm_child(xml_node, "d", level = c("b", "c"))
#'
#' @export
xml_rm_child <- function(xml_node, xml_child, level, which = 0, pointer = FALSE, ...) {

  if (missing(xml_node))
    stop("need xml_node")

  if (missing(xml_child))
    stop("need xml_child")

  if (!inherits(xml_node, "pugi_xml")) xml_node <- read_xml(xml_node, ...)
  assert_class(xml_child, "character")

  which <- which - 1

  if (missing(level)) {
    z <- xml_remove_child1(xml_node, xml_child, which, pointer)
  } else {

    if (length(level) == 1)
      z <- xml_remove_child2(xml_node, xml_child, level[[1]], which, pointer)

    if (length(level) == 2)
      z <- xml_remove_child3(xml_node, xml_child, level[[1]], level[[2]], which, pointer)

  }

  return(z)
}
