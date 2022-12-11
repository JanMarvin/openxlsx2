workbook_set_header_footer <- function(
    self,
    private,
    sheet       = current_sheet(),
    header      = NULL,
    footer      = NULL,
    evenHeader  = NULL,
    evenFooter  = NULL,
    firstHeader = NULL,
    firstFooter = NULL
) {
  sheet <- private$get_sheet_index(sheet)

  if (!is.null(header) && length(header) != 3) {
    stop("header must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(footer) && length(footer) != 3) {
    stop("footer must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(evenHeader) && length(evenHeader) != 3) {
    stop("evenHeader must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(evenFooter) && length(evenFooter) != 3) {
    stop("evenFooter must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(firstHeader) && length(firstHeader) != 3) {
    stop("firstHeader must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(firstFooter) && length(firstFooter) != 3) {
    stop("firstFooter must have length 3 where elements correspond to positions: left, center, right.")
  }

  # TODO this could probably be moved to the hf assignment
  oddHeader   <- headerFooterSub(header)
  oddFooter   <- headerFooterSub(footer)
  evenHeader  <- headerFooterSub(evenHeader)
  evenFooter  <- headerFooterSub(evenFooter)
  firstHeader <- headerFooterSub(firstHeader)
  firstFooter <- headerFooterSub(firstFooter)

  hf <- list(
    oddHeader = naToNULLList(oddHeader),
    oddFooter = naToNULLList(oddFooter),
    evenHeader = naToNULLList(evenHeader),
    evenFooter = naToNULLList(evenFooter),
    firstHeader = naToNULLList(firstHeader),
    firstFooter = naToNULLList(firstFooter)
  )

  if (all(lengths(hf) == 0)) {
    hf <- NULL
  }

  self$worksheets[[sheet]]$headerFooter <- hf
  invisible(self)
}
