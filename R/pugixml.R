

#' read xml file
#' @param x somthing to read
#' @param pointer return pointer?
#' @export
read_xml <- function(xml, pointer = TRUE)  {
  
  z <- NULL 
  
  if (pointer) {
    z <- readXMLPtr(xml)
  }
  else {
    z <- readXML(xml)
  }
  
  z
}


#' xml_node
#' @param x something xml
#' @param level1 to please check
#' @param level2 to please check
#' @param level3 to please check
#' @param level4 to please check
#' @param level5 to please check
#' @export
xml_node <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL, level4 = NULL, level5 = NULL, level6 = NULL) {
  
  lvl <- c(level1, level2, level3, level4, level5, level6)
  lvl <- lvl[!is.null(lvl)]
  if (!all(is.character(lvl)))
    stop("levels must be character vectors")
  
  z <- NULL
  
  if (class(xml) == "pugi_xml") {
    if (length(lvl) == 1) z <- getXMLXPtr1(xml, level1)
    if (length(lvl) == 2) z <- getXMLXPtr2(xml, level1, level2)
    if (length(lvl) == 3) z <- getXMLXPtr3(xml, level1, level2, level3)
    if (length(lvl) == 4) z <- getXMLXPtr4(xml, level1, level2, level3, level4)
    if (length(lvl) == 5) z <- getXMLXPtr5(xml, level1, level2, level3, level4, level5)
  } else {
    if (length(lvl) == 1) z <- getXML1(xml, level1)
    if (length(lvl) == 2) z <- getXML2(xml, level1, level2)
    if (length(lvl) == 3) z <- getXML3(xml, level1, level2, level3)
  }
  
  z
}

#' xml_value
#' @param x something xml
#' @param level1 to please check
#' @param level2 to please check
#' @param level3 to please check
#' @param level4 to please check
#' @param level5 to please check
#' @export
xml_value <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL, level4 = NULL, level5 = NULL, level6 = NULL) {
  
  lvl <- c(level1, level2, level3, level4, level5, level6)
  lvl <- lvl[!is.null(lvl)]
  if (!all(is.character(lvl)))
    stop("levels must be character vectors")
  
  z <- NULL
  
  if (class(xml) == "pugi_xml") {
    if (length(lvl) == 2) z <- getXMLXPtr2val(xml, level1, level2)
    if (length(lvl) == 3) z <- getXMLXPtr3val(xml, level1, level2, level3)
    if (length(lvl) == 4) z <- getXMLXPtr4val(xml, level1, level2, level3, level4)
    if (length(lvl) == 5) z <- getXMLXPtr5val(xml, level1, level2, level3, level4, level5)
  } else {
    if (length(lvl) == 1) z <- getXML1val(xml, level1)
    if (length(lvl) == 2) z <- getXML2val(xml, level1, level2)
    if (length(lvl) == 3) z <- getXML3val(xml, level1, level2, level3)
  }
  
  z
}

#' xml_attribute
#' @param x something xml
#' @param level1 to please check
#' @param level2 to please check
#' @param level3 to please check
#' @param level4 to please check
#' @param level5 to please check
#' @export
xml_attribute <- function(xml, level1 = NULL, level2 = NULL, level3 = NULL, level4 = NULL, level5 = NULL, level6 = NULL) {
  
  lvl <- c(level1, level2, level3, level4, level5, level6)
  lvl <- lvl[!is.null(lvl)]
  if (!all(is.character(lvl)))
    stop("levels must be character vectors")
  
  z <- NULL
  
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

#' @method print pugi_xml
#' @param x somthing to print
#' @param raw print as raw text
#' @param ... to please check
#' @export
print.pugi_xml <- function(x, raw = FALSE, ...) {
  cat(printXPtr(x, raw))
}