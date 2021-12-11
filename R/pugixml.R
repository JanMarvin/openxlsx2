

#' read xml file
#' @param xml somthing to read character string or file
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
#' @export
read_xml <- function(xml, pointer = TRUE)  {

  z <- NULL

  isfile = FALSE
  if (file.exists(xml))
    isfile <- TRUE

  if (pointer) {
    z <- readXMLPtr(xml, isfile)
  }
  else {
    z <- readXML(xml, isfile)
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
  } else { # should not happen?
    xml <- read_xml(xml, pointer = FALSE)
    if (length(lvl) == 1) z <- getXML1(xml, level1)
    if (length(lvl) == 2) z <- getXML2(xml, level1, level2)
    if (length(lvl) == 3) z <- getXML3(xml, level1, level2, level3)
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
  } else {
    xml <- read_xml(xml, pointer = FALSE)
    if (length(lvl) == 1) z <- getXML1val(xml, level1)
    if (length(lvl) == 2) z <- getXML2val(xml, level1, level2)
    if (length(lvl) == 3) z <- getXML3val(xml, level1, level2, level3)
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
  } else {
    warning("nothing available")
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
  cat(printXPtr(x, raw))
}
