#' R6 class for a Workbook Comments
#'
#' A comment
#'
#' @export
wbComment <- R6::R6Class(
  "wbComment",

  public = list(
    #' @field text Comment text
    text = character(),

    #' @field author The comment author
    author = character(),

    #' @field style A style for the comment
    style = character(),

    #' @field visible `logical`, if `FALSE` is not visible
    visible = TRUE,

    # TODO what unit is width/height?
    #' @field width Width of the comment in ... units
    width = 2,

    #' @field height Height of comment in ... units
    height = 4,

    #' @description
    #' Creates a new `wbComment` object
    #' @param text Comment text
    #' @param author The comment author
    #' @param style A style for the comment
    #' @param visible `logical`, if `FALSE` is not visible
    #' @param width Width of the comment in ... units
    #' @param height Height of comment in ... units
    #' @return a `wbComment` object
    initialize = function(text, author, style, visible = TRUE, width = 2, height = 4) {
      # TODO this needs the validations that the comment wrappers have
      self$text <- text
      self$author <- author
      self$style <- style
      self$visible <- visible
      self$width <- width
      self$height <- height
      invisible(self)
    },

    #' @description
    #' Prints the object
    #' @returns The `wbComment` object, invisibly; called for its side effects
    print = function() {
      showText <- c(
        sprintf("Author: %s\n", self$author),
        sprintf("Text:\n %s\n\n", paste(self$text, collapse = ""))
      )

      s <- self$style

      if (!inherits(self$style, "character"))
        stop("style must be a character string: something like <font>...</font>")

      styleShow <- "Style:\n"
      for (i in seq_along(s)) {
        styleShow <- c(
          styleShow,
          sprintf("Font name: %s\n", unname(unlist(xml_attr(s[[i]], "font", "name")))), ## Font name
          sprintf("Font size: %s\n", unname(unlist(xml_attr(s[[i]], "font", "sz")))), ## Font size
          sprintf("Font colour: %s\n", gsub("^FF", "#", unname(unlist(xml_attr(s[[i]], "font", "color"))))), ## Font colour
          "\n\n"
        )
      }

      cat(showText, styleShow, sep = "")
      invisible(self)
    }
  )
)


# wrappers ----------------------------------------------------------------

# TODO create_comment() should leverage wbComment$new() more
# TODO write_comment() should leverage wbWorkbook$addComment() more
# TODO remove_comment() should leverage wbWorkbook$remove_comment() more

#' @name create_comment
#' @title Create, write and remove comments
#' @description The comment functions (create, write and remove) allow the
#' modification of comments. In newer Excels they are called notes, while they
#' are called comments in openxml. Modification of what Excel now calls comment
#' (openxml calls them threadedComments) is not yet possible
#' @param text Comment text. Character vector.
#' @param author Author of comment. Character vector of length 1
#' @param style A Style object or list of style objects the same length as comment vector.
#' @param visible TRUE or FALSE. Is comment visible.
#' @param width Textbox integer width in number of cells
#' @param height Textbox integer height in number of cells
#' @export
#' @rdname comment
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#'
#' # write comment without author
#' c1 <- create_comment(text = "this is a comment", author = "")
#' write_comment(wb, 1, col = "B", row = 10, comment = c1)
#'
#' # Write another comment with author information
#' c2 <- create_comment(text = "this is another comment", author = "Marco Polo")
#' write_comment(wb, 1, col = "C", row = 10, comment = c2)
#'
#' # write a styled comment with system author
#' s1 <- create_font(b = "true", color = c(rgb = "FFFF0000"), sz = "12")
#' s2 <- create_font(color = c(rgb = "FF000000"), sz = "9")
#' c3 <- create_comment(text = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
#'
#' write_comment(wb, 1, col = 6, row = 3, comment = c3)
#'
#' # remove the first comment
#' remove_comment(wb, 1, col = "B", row = 10)
#' \dontrun{
#' wb_save(wb, path = "create_commentExample.xlsx", overwrite = TRUE)
#' }
create_comment <- function(text,
  author = Sys.info()[["user"]],
  style = NULL,
  visible = TRUE,
  width = 2,
  height = 4) {

  # TODO move this to wbComment$new(); this could then be replaced with
  # wb_comment()

  assert_class(author, "character")
  assert_class(text, "character")
  assert_class(width, "numeric")
  assert_class(height, "numeric")
  assert_class(visible, "logical")

  if (length(visible) > 1) stop("visible must be a single logical")
  if (length(author) > 1) stop("author) must be a single character")

  width <- round(width)
  height <- round(height)

  author <- author[1]
  visible <- visible[1]

  if (is.null(style)) {
    style <- create_font()
  }

  author <- replaceIllegalCharacters(author)
  text <- replaceIllegalCharacters(text)


  invisible(wbComment$new(text = text, author = author, style = style, visible = visible, width = width[1], height = height[1]))
}


#' @name write_comment
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @param col Column a column number of letter. For `remove_comment` this can be a range.
#' @param row A row number. For `remove_comment` this can be a range.
#' @param comment A Comment object. See [create_comment()].
#' @param xy An alternative to specifying `col` and
#' `row` individually.  A vector of the form
#' `c(col, row)`.
#' @rdname comment
#' @export
write_comment <- function(wb, sheet, col, row, comment, xy = NULL) {
  # TODO add as method: wbWorkbook$addComment(); add param for replace?
  assert_workbook(wb)
  assert_comment(comment)

  # rPr <- wb$createFontNode(NULL) # TODO not yet possible
  if (is.null(comment$style)) {
    rPr <- create_font()
    } else {
    rPr <- comment$style
  }

  rPr <- gsub("font>", "rPr>", rPr)
  sheet <- wb_validate_sheet(wb, sheet)

  ## All input conversions/validations
  if (!is.null(xy)) {
    if (length(xy) != 2) {
      stop("xy parameter must have length 2")
    }
    col <- xy[[1]]
    row <- xy[[2]]
  }

  if (!is.numeric(col)) {
    col <- col2int(col)
  }

  ref <- paste0(int2col(col), row)

  comment_list <- list(
    "ref" = ref,
    "author" = comment$author,
    "comment" = comment$text,
    "style" = rPr,
    "clientData" = genClientData(col, row, visible = comment$visible, height = comment$height, width = comment$width)
  )

  fn <- sprintf("comments%s.xml", sheet)

  if (!any(grepl(fn, wb$Content_Types))) {
    wb$Content_Types <- c(
      wb$Content_Types,
      sprintf(
        '<Override PartName="/xl/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>',
        fn
      )
    )
  }

  wb$comments[[sheet]] <- c(wb$comments[[sheet]], list(comment_list))

  relship <- rbindlist(xml_attr(wb$worksheets_rels[[sheet]], "Relationship"))
  relship$typ <- basename(relship$Type)
  relship$id  <- as.integer(gsub("\\D+", "", relship$Id))

  next_rid <- max(relship$id) + 1
  if (any(relship$typ == "comments"))
    next_rid <- relship$id[relship$typ == "comments"]

  # unique? keep prev legacyDrawing?
  #self$worksheets[[i]]$legacyDrawing <- '<legacyDrawing r:id="rId2"/>'
  # TODO hardcoded 2. Marvin fears that this is not good enough
  wb$worksheets[[sheet]]$legacyDrawing <- sprintf('<legacyDrawing r:id="rId%s"/>', 2)


  wb$worksheets_rels[[sheet]] <- unique(c(
    wb$worksheets_rels[[sheet]],
    sprintf(
      '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../%s"/>',
      next_rid,
      fn
    )
  ))

  invisible(wb)
}


#' @name remove_comment
#' @param gridExpand If `TRUE`, all data in rectangle min(rows):max(rows) X min(cols):max(cols)
#' will be removed.
#' @rdname comment
#' @export
remove_comment <- function(wb, sheet, col, row, gridExpand = TRUE) {
  # TODO add as method; wbWorkbook$remove_comment()
  assert_workbook(wb)

  sheet <- wb_validate_sheet(wb, sheet)

  # col2int checks for numeric
  col <- col2int(col)
  row <- as.integer(row)

  ## rows and cols need to be the same length
  if (gridExpand) {
    combs <- expand.grid(row, col)
    row <- combs[, 1]
    col <- combs[, 2]
  }

  if (length(row) != length(col)) {
    stop("Length of rows and cols must be equal.")
  }

  comb <- paste0(int2col(col), row)
  toKeep <- !sapply(wb$comments[[sheet]], "[[", "ref") %in% comb

  wb$comments[[sheet]] <- wb$comments[[sheet]][toKeep]
}

wb_comment <- function(text = character(), author = character(), style = character()) {
  wbComment$new(text = text, author = author, style = style)
}
