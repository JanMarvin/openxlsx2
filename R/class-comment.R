
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

    #' @field style A style (class `wbStyle`) for the comment (?)
    style = NULL,

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
    #' @param style A style (class `wbStyle`) for the comment (?)
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

      # TODO style should probably always be a list?
      # TODO would style be a style object?
      s <- if (inherits(self$style, "list")) {
        self$style
      }  else  {
        list(self$style)
      }

      styleShow <- "Style:\n"
      for (i in seq_along(s)) {
        styleShow <- c(
          styleShow,
          sprintf("Font name: %s\n", s[[i]]$fontName[[1]]), ## Font name
          sprintf("Font size: %s\n", s[[i]]$fontSize[[1]]), ## Font size
          sprintf("Font colour: %s\n", gsub("^FF", "#", s[[i]]$fontColour[[1]])), ## Font colour
          ## Font decoration
          if (length(s[[i]]$fontDecoration)) {
            sprintf("Font decoration: %s\n", paste(s[[i]]$fontDecoration, collapse = ", "))
          },
          "\n\n"
        )
      }

      cat(showText, styleShow, sep = "")
      invisible(self)
    }
  )
)


# wrappers ----------------------------------------------------------------

# TODO createComment() should leverage wbwbComment$new() more
# TODO writeComment() should leverage wbWorkbook$addComment() more
# TODO removeComment() should leverage wbWorkbook$removeComment() more

#' @name createComment
#' @title create a Comment object
#' @description Create a cell Comment object to pass to writeComment()
#' @param comment Comment text. Character vector.
#' @param author Author of comment. Character vector of length 1
#' @param style A Style object or list of style objects the same length as comment vector. See [createStyle()].
#' @param visible TRUE or FALSE. Is comment visible.
#' @param width Textbox integer width in number of cells
#' @param height Textbox integer height in number of cells
#' @export
#' @seealso [writeComment()]
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#'
#' c1 <- createComment(comment = "this is comment")
#' writeComment(wb, 1, col = "B", row = 10, comment = c1)
#'
#' s1 <- createStyle(fontSize = 12, fontColour = "red", textDecoration = "bold")
#' s2 <- createStyle(fontSize = 9, fontColour = "black")
#'
#' c2 <- createComment(comment = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
#' c2
#'
#' writeComment(wb, 1, col = 6, row = 3, comment = c2)
#' \dontrun{
#' saveWorkbook(wb, file = "createCommentExample.xlsx", overwrite = TRUE)
#' }
createComment <- function(comment,
  author = Sys.getenv("USERNAME"),
  style = NULL,
  visible = TRUE,
  width = 2,
  height = 4) {

  # TODO move this to wbComment$new(); this could then be replaced with
  # wb_comment()

  assert_class(author, "character")
  assert_class(comment, "character")
  assert_class(width, "numeric")
  assert_class(height, "numeric")
  assert_class(visible, "logical")

  width <- round(width)
  height <- round(height)

  author <- author[1]
  visible <- visible[1]

  if (is.null(style)) {
    style <- createStyle(fontName = "Tahoma", fontSize = 9, fontColour = "black")
  }

  author <- replaceIllegalCharacters(author)
  comment <- replaceIllegalCharacters(comment)


  invisible(wbComment$new(text = comment, author = author, style = style, visible = visible, width = width[1], height = height[1]))
}



#' @name writeComment
#' @title write a cell comment
#' @description Write a Comment object to a worksheet
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @param col Column a column number of letter
#' @param row A row number.
#' @param comment A Comment object. See [createComment()].
#' @param xy An alternative to specifying `col` and
#' `row` individually.  A vector of the form
#' `c(col, row)`.
#' @export
#' @seealso [createComment()]
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#'
#' c1 <- createComment(comment = "this is comment")
#' writeComment(wb, 1, col = "B", row = 10, comment = c1)
#'
#' s1 <- createStyle(fontSize = 12, fontColour = "red", textDecoration = "bold")
#' s2 <- createStyle(fontSize = 9, fontColour = "black")
#'
#' c2 <- createComment(comment = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
#' c2
#'
#' writeComment(wb, 1, col = 6, row = 3, comment = c2)
#' \dontrun{
#' saveWorkbook(wb, file = "writeCommentExample.xlsx", overwrite = TRUE)
#' }
writeComment <- function(wb, sheet, col, row, comment, xy = NULL) {
  # TODO add as method: wbWorkbook$addComment(); add param for replace?
  assert_workbook(wb)
  assert_comment(comment)

  # if (length(comment$style) == 1) {
  #   rPr <- wb$createFontNode(comment$style)
  # } else {
  #   rPr <- sapply(comment$style, function(x) wb$createFontNode(x))
  # }
  assert_comment(comment)
  rPr <- wb$createFontNode(comment$style)

  rPr <- gsub("font>", "rPr>", rPr)
  sheet <- wb$validateSheet(sheet)

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

  wb$comments[[sheet]] <- c(wb$comments[[sheet]], list(comment_list))

  invisible(wb)
}



#' @name removeComment
#' @title Remove a comment from a cell
#' @description Remove a cell comment from a worksheet
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @param cols Columns to delete comments from
#' @param rows Rows to delete comments from
#' @param gridExpand If `TRUE`, all data in rectangle min(rows):max(rows) X min(cols):max(cols)
#' will be removed.
#' @export
#' @seealso [createComment()]
#' @seealso [writeComment()]
removeComment <- function(wb, sheet, cols, rows, gridExpand = TRUE) {
  # TODO add as method; wbWorkbook$removeComment()
  assert_workbook(wb)

  sheet <- wb$validateSheet(sheet)
  cols <- col2int(cols)
  rows <- as.integer(rows)

  ## rows and cols need to be the same length
  if (gridExpand) {
    combs <- expand.grid(rows, cols)
    rows <- combs[, 1]
    cols <- combs[, 2]
  }

  if (length(rows) != length(cols)) {
    stop("Length of rows and cols must be equal.")
  }

  comb <- paste0(int2col(cols), rows)
  toKeep <- !sapply(wb$comments[[sheet]], "[[", "ref") %in% comb

  wb$comments[[sheet]] <- wb$comments[[sheet]][toKeep]
}

wb_comment <- function(text = character(), author = character(), style = wb_style()) {
  wbComment$new(text = text, author = author, style = style)
}
