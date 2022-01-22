

#' read xml file
#' @param xml something to read character string or file
#' @param declaration should the declaration be imported
#' @param escapes bool if characters like "&" should be escaped. The default is
#' no escapes. Assuming that the input already provides valid information.
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
read_xml <- function(xml, pointer = TRUE, escapes = FALSE, declaration = FALSE) {

  z <- NULL

  isfile = FALSE
  if (length(xml) == 1 && file.exists(xml))
    isfile <- TRUE

  if (!isfile)
    xml <- paste0(xml, collapse = "")

  if (pointer) {
    z <- readXMLPtr(xml, isfile, escapes, declaration)
  }
  else {
    z <- readXML(xml, isfile, escapes, declaration)
  }

  z
}


#' xml_node
#' @param xml something xml
#' @param level1 to please check
#' @param level2 to please check
#' @param level3 to please check
#' @param level4 to please check
#' @param level5 to please check
#' @param level6 to please check
#' @details This function returns XML nodes as used in openxlsx2. In theory they
#' could be returned as pointers as well, but this has not yet been implemented.
#' @examples
#'   x <- read_xml("<a><b/></a>")
#'   # return a
#'   xml_node(x, "a")
#'   # return b. requires the path to the node
#'   xml_node(x, "a", "b")
#' @export
xml_node <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL, level4 = NULL, level5 = NULL, level6 = NULL) {

  lvl <- c(level1, level2, level3, level4, level5, level6)
  lvl <- lvl[!is.null(lvl)]
  if (!all(is.character(lvl)))
    stop("levels must be character vectors")

  z <- NULL

  if(class(xml) != "pugi_xml")
    xml <- read_xml(xml)


  if (class(xml) == "pugi_xml") {
    if (length(lvl) == 1) z <- getXMLXPtr1(xml, level1)
    if (length(lvl) == 2) z <- getXMLXPtr2(xml, level1, level2)
    if (length(lvl) == 3) z <- getXMLXPtr3(xml, level1, level2, level3)
    if (length(lvl) == 3) if (level2 == "*") z <- unkgetXMLXPtr3(xml, level1, level3)
    if (length(lvl) == 4) z <- getXMLXPtr4(xml, level1, level2, level3, level4)
    if (length(lvl) == 5) z <- getXMLXPtr5(xml, level1, level2, level3, level4, level5)
  }

  z
}

#' xml_value
#' @param xml something xml
#' @param level1 to please check
#' @param level2 to please check
#' @param level3 to please check
#' @param level4 to please check
#' @param level5 to please check
#' @param level6 to please check
#' @examples#'
#'   x <- read_xml("<a>1</a>")
#'   xml_value(x, "a")
#'
#'   x <- read_xml("<a><b r=\"1\">2</b></a>")
#'   xml_value(x, "a", "b")
#' @export
xml_value <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL, level4 = NULL, level5 = NULL, level6 = NULL) {

  lvl <- c(level1, level2, level3, level4, level5, level6)
  lvl <- lvl[!is.null(lvl)]
  if (!all(is.character(lvl)))
    stop("levels must be character vectors")

  z <- NULL

  if(class(xml) != "pugi_xml")
    xml <- read_xml(xml)

  if (class(xml) == "pugi_xml") {
    if (length(lvl) == 1) z <- getXMLXPtr1val(xml, level1)
    if (length(lvl) == 2) z <- getXMLXPtr2val(xml, level1, level2)
    if (length(lvl) == 3) z <- getXMLXPtr3val(xml, level1, level2, level3)
    if (length(lvl) == 4) z <- getXMLXPtr4val(xml, level1, level2, level3, level4)
    if (length(lvl) == 5) z <- getXMLXPtr5val(xml, level1, level2, level3, level4, level5)
  }

  z
}

#' xml_attribute
#' @param xml something xml
#' @param level1 to please check
#' @param level2 to please check
#' @param level3 to please check
#' @param level4 to please check
#' @param level5 to please check
#' @param level6 to please check
#' @examples
#'
#'   x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
#'   xml_attribute(x, "a")
#'
#'   x <- read_xml("<a><b r=\"1\">2</b></a>")
#'   xml_attribute(x, "a", "b")
#'   x <- read_xml("<a a=\"1\" b=\"2\">1</a>")
#'   xml_attribute(x, "a")
#'
#'   x <- read_xml("<b><a a=\"1\" b=\"2\"/></b>")
#'   xml_attribute(x, "b", "a")
#' @export
xml_attribute <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL, level4 = NULL, level5 = NULL, level6 = NULL) {

  lvl <- c(level1, level2, level3, level4, level5, level6)
  lvl <- lvl[!is.null(lvl)]
  if (!all(is.character(lvl)))
    stop("levels must be character vectors")

  z <- NULL

  if(class(xml) != "pugi_xml")
    xml <- read_xml(xml)

  if (class(xml) == "pugi_xml") {
    if (length(lvl) == 1) z <- getXMLXPtr1attr(xml, level1)
    if (length(lvl) == 2) z <- getXMLXPtr2attr(xml, level1, level2)
    if (length(lvl) == 3) z <- getXMLXPtr3attr(xml, level1, level2, level3)
    if (length(lvl) == 4) z <- getXMLXPtr4attr(xml, level1, level2, level3, level4)
    if (length(lvl) == 5) z <- getXMLXPtr5attr(xml, level1, level2, level3, level4, level5)
  }

  z
}

#' print pugi_xml
#' @method print pugi_xml
#' @param x somthing to print
#' @param raw print as raw text
#' @param ... to please check
#' @examples
#'   # a pointer
#'   x <- read_xml("<a><b/></a>")
#'   print(x)
#'   print(x, raw = TRUE)
#' @export
print.pugi_xml <- function(x, raw = FALSE, ...) {

  escapes <- attr(x, "escapes")

  cat(printXPtr(x, !escapes, raw))
  if (raw) cat("\n")
}

#' loads character string to pugixml and returns an externalptr
#' @details
#' might be usefull for larger documents where single nodes are shortened
#' and otherwise the full tree has to be reimported. unsure where we have
#' such a case.
#' is usefull, for printing nodes from a larger tree, that have been exported
#' as characters (at some point in time we have to convert the xml to R)
#' @param x input as xml
#' @examples
#' \dontrun{
#' tmp_xlsx <- tempdir()
#' xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")
#' unzip(xlsxFile, exdir = tmp_xlsx)
#'
#' wb <- loadWorkbook(xlsxFile)
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
#' }
#' @export
as_xml <- function(x) {
  read_xml(paste(x, collapse = ""))
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
write_file <- function(head = "", body = "", tail = "", fl = "", escapes = FALSE) {
  xml_content <- paste0(head, body, tail, collapse = "")
  write_xml_file(xml_content = xml_content, fl = fl, escapes)
}

#' append xml child to node
#' @param xml_node xml_node
#' @param xml_child xml_child
#' @param pointer pointer
#' @examples
#' xml_node <- "<node><child1/><child2/></node>"
#' xml_child <- "<new_child/>"
#'
#' xml_add_child(xml_node, xml_child)
#' @export
xml_add_child <- function(xml_node, xml_child, pointer = FALSE) {

  if (missing(xml_node))
    stop("need xml_node")

  if (missing(xml_child))
    stop("need xml_child")

  xml_node <- read_xml(xml_node)
  xml_child <- read_xml(xml_child)

  z <- xml_append_child(xml_node, xml_child, pointer)

  return(z)
}
