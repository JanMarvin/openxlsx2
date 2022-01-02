#' @name convertToDate
#' @title Convert from excel date number to R Date type
#' @description Convert from excel date number to R Date type
#' @param x A vector of integers
#' @param origin date. Default value is for Windows Excel 2010
#' @param ... additional parameters passed to as.Date()
#' @details Excel stores dates as number of days from some origin day
#' @seealso \code{\link{writeData}}
#' @export
#' @examples
#' ## 2014 April 21st to 25th
#' convertToDate(c(41750, 41751, 41752, 41753, 41754, NA))
#' convertToDate(c(41750.2, 41751.99, NA, 41753))
convertToDate <- function(x, origin = "1900-01-01", ...) {
  x <- as.numeric(x)
  notNa <- !is.na(x)
  earlyDate <- x < 60
  if (origin == "1900-01-01") {
    x[notNa] <- x[notNa] - 2
    x[earlyDate & notNa] <- x[earlyDate & notNa] + 1
  }

  return(as.Date(x, origin = origin, ...))
}


#' @name convertToDateTime
#' @title Convert from excel time number to R POSIXct type.
#' @description Convert from excel time number to R POSIXct type.
#' @param x A numeric vector
#' @param origin date. Default value is for Windows Excel 2010
#' @param ... Additional parameters passed to as.POSIXct
#' @details Excel stores dates as number of days from some origin date
#' @export
#' @examples
#' ## 2014-07-01, 2014-06-30, 2014-06-29
#' x <- c(41821.8127314815, 41820.8127314815, NA, 41819, NaN)
#' convertToDateTime(x)
#' convertToDateTime(x, tz = "Australia/Perth")
#' convertToDateTime(x, tz = "UTC")
convertToDateTime <- function(x, origin = "1900-01-01", ...) {
  sci_pen <- getOption("scipen")
  options("scipen" = 10000)
  on.exit(options("scipen" = sci_pen), add = TRUE)

  x <- as.numeric(x)
  date <- convertToDate(x, origin)

  x <- x * 86400
  rem <- x %% 86400

  hours <- as.integer(floor(rem / 3600))
  minutes_fraction <- rem %% 3600
  minutes_whole <- as.integer(floor(minutes_fraction / 60))
  secs <- minutes_fraction %% 60

  y <- sprintf("%02d:%02d:%06.3f", hours, minutes_whole, secs)
  notNA <- !is.na(x)
  date_time <- rep(NA, length(x))
  date_time[notNA] <- as.POSIXct(paste(date[notNA], y[notNA]), ...)

  date_time <- .POSIXct(date_time)

  return(date_time)
}



#' @name getDateOrigin
#' @title Get the date origin an xlsx file is using
#' @description Return the date origin used internally by an xlsx or xlsm file
#' @author Alexander Walker
#' @param xlsxFile An xlsx or xlsm file.
#' @details Excel stores dates as the number of days from either 1904-01-01 or 1900-01-01. This function
#' checks the date origin being used in an Excel file and returns is so it can be used in \code{\link{convertToDate}}
#' @return One of "1900-01-01" or "1904-01-01".
#' @seealso \code{\link{convertToDate}}
#' @examples
#'
#' ## create a file with some dates
#' \dontrun{
#' write.xlsx(as.Date("2015-01-10") - (0:4), file = "getDateOriginExample.xlsx")
#' m <- read.xlsx("getDateOriginExample.xlsx")
#'
#' ## convert to dates
#' do <- getDateOrigin(system.file("extdata", "readTest.xlsx", package = "openxlsx2"))
#' convertToDate(m[[1]], do)
#' }
#' @export
getDateOrigin <- function(xlsxFile) {
  xlsxFile <- getFile(xlsxFile)
  if (!file.exists(xlsxFile)) {
    stop("File does not exist.")
  }

  if (grepl("\\.xls$|\\.xlm$", xlsxFile)) {
    stop("openxlsx can not read .xls or .xlm files!")
  }

  ## create temp dir and unzip
  xmlDir <- file.path(tempdir(), "_excelXMLRead")
  xmlFiles <- unzip(xlsxFile, exdir = xmlDir)

  on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)

  workbook <- grep("workbook.xml$", xmlFiles, perl = TRUE, value = TRUE)
  workbook <- paste(unlist(readUTF8(workbook)), collapse = "")

  if (grepl('date1904="1"|date1904="true"', workbook, ignore.case = TRUE)) {
    origin <- "1904-01-01"
  } else {
    origin <- "1900-01-01"
  }

  return(origin)
}
