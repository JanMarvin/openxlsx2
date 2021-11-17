


genBaseColStyle <- function(cc) {
  colStyle <- createStyle()
  specialFormat <- TRUE

  if ("date" %in% cc) {
    colStyle <- createStyle("numFmt" = "date")
  } else if (any(c("posixlt", "posixct", "posixt") %in% cc)) {
    colStyle <- createStyle("numFmt" = "longdate")
  } else if ("currency" %in% cc) {
    colStyle$numFmt <- list("numFmtId" = "164", "formatCode" = "&quot;$&quot;#,##0.00")
  } else if ("accounting" %in% cc) {
    colStyle$numFmt <- list("numFmtId" = "44")
  } else if ("hyperlink" %in% cc) {
    colStyle$fontColour <- list("theme" = "10")
  } else if ("percentage" %in% cc) {
    colStyle$numFmt <- list("numFmtId" = "10")
  } else if ("scientific" %in% cc) {
    colStyle$numFmt <- list("numFmtId" = "11")
  } else if ("3" %in% cc | "comma" %in% cc) {
    colStyle$numFmt <- list("numFmtId" = "3")
  } else if ("numeric" %in% cc & !grepl("[^0\\.,#\\$\\* %]", getOption("openxlsx.numFmt", "GENERAL"))) {
    colStyle$numFmt <- list("numFmtId" = 9999, "formatCode" = getOption("openxlsx.numFmt"))
  } else {
    colStyle$numFmt <- list(numFmtId = "0")
    specialFormat <- FALSE
  }

  list(
    "style" = colStyle,
    "specialFormat" = specialFormat
  )
}
