#' R6 class for a Workbook Comments
#'
#' A comment
#'
#' @noRd
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
    initialize = function(text, author, style, visible, width, height) {
      self$text    <- text
      self$author  <- author
      self$style   <- style
      self$visible <- visible
      self$width   <- width
      self$height  <- height
      invisible(self)
    },

    #' @description
    #' Prints the object
    #' @returns The `wbComment` object, invisibly; called for its side effects
    print = function() {
      showText <- c(
        sprintf("Author: %s\n", self$author),
        sprintf("Text:\n %s\n\n", paste(as.character(self$text), collapse = ""))
      )

      s <- self$style

      if (!inherits(self$style, "character"))
        stop("style must be a character string: something like <font>...</font>")

      styleShow <- "Style:\n"
      for (i in seq_along(s)) {
        styleShow <- c(
          styleShow,
          sprintf("Font name: %s\n", unlist(xml_attr(s[[i]], "font", "name"), use.names = FALSE)), ## Font name
          sprintf("Font size: %s\n", unlist(xml_attr(s[[i]], "font", "sz"), use.names = FALSE)), ## Font size
          sprintf("Font color: %s\n", gsub("^FF", "#", unlist(xml_attr(s[[i]], "font", "color"), use.names = FALSE))), ## Font color
          "\n\n"
        )
      }

      cat(showText, styleShow, sep = "")
      invisible(self)
    }
  )
)
# Comment creation wrappers ----------------------------------------------------------------

# TODO wb_comment() should leverage wbComment$new() more

#' Helper to create a comment object
#'
#' Creates a `wbComment` object. Use with [wb_add_comment()] to add to a worksheet location.
#'
#' @param text Comment text. Character vector. or a [fmt_txt()] string.
#' @param author Author of comment. A string. By default, will look at `options("openxlsx2.creator")`.
#'   Otherwise, will check the system username.
#' @param style A Style object or list of style objects the same length as comment vector.
#' @param visible Is comment visible? Default: `FALSE`.
#' @param width Textbox integer width in number of cells
#' @param height Textbox integer height in number of cells
#'
#' @return A `wbComment` object
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#'
#' # write comment without author
#' c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
#' wb$add_comment(dims = "B10", comment = c1)
#'
#' # Write another comment with author information
#' c2 <- wb_comment(text = "this is another comment", author = "Marco Polo")
#' wb$add_comment(sheet = 1, dims = "C10", comment = c2)
#'
#' # write a styled comment with system author
#' s1 <- create_font(b = "true", color = wb_color(hex = "FFFF0000"), sz = "12")
#' s2 <- create_font(color = wb_color(hex = "FF000000"), sz = "9")
#' c3 <- wb_comment(text = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
#'
#' wb$add_comment(sheet = 1, dims = wb_dims(3, 6), comment = c3)
wb_comment <- function(
    text    = NULL,
    style   = NULL,
    visible = FALSE,
    author  = getOption("openxlsx2.creator"),
    width   = 2,
    height  = 4
) {
  # Code copied from the wbWorkbook
  author <- author %||% Sys.getenv("USERNAME", unset = Sys.getenv("USER"))
  text <- text %||% ""
  assert_class(author, "character")
  assert_class(text, "character")
  assert_class(width, c("numeric", "integer"))
  assert_class(height, c("numeric", "integer"))
  assert_class(visible, "logical")

  if (length(visible) > 1) stop("visible must be a single logical")
  if (length(author) > 1)  stop("author) must be a single character")
  if (length(width) > 1)   stop("width must be a single integer")
  if (length(height) > 1)  stop("height must be a single integer")

  width  <- round(width)
  height <- round(height)

  if (is.null(style)) {
    style <- create_font()
  }

  # if text was created using fmt_txt()
  if (inherits(text, "fmt_txt")) {
    text  <- text
    style <- ""
  } else {
    text <- replace_legal_chars(text)
  }

  author <- replace_legal_chars(author)
  if (author != "") {
    # if author is provided, we write additional lines with the author name as well as an empty line
    text <- c(paste0(author, ":"), "\n", text)
    style <- c(
      # default node consist of these two styles for the author name and the empty line.
      # values are default in MS365
      '<rPr><b/><sz val=\"10\"/><color rgb=\"FF000000\"/><rFont val=\"Tahoma\"/><family val=\"2\"/></rPr>',
      '<rPr><sz val=\"10\"/><color rgb=\"FF000000\"/><rFont val=\"Tahoma\"/><family val=\"2\"/></rPr>',
      style
    )
  }

  invisible(wbComment$new(text = text, author = author, style = style, width = width, height = height, visible = visible))
}

do_write_comment <- function(
    wb,
    sheet,
    col     = NULL,
    row     = NULL,
    comment,
    dims    = rowcol_to_dim(row, col),
    color   = NULL,
    file    = NULL
  ) {

  # TODO add as method: wbWorkbook$addComment(); add param for replace?
  assert_workbook(wb)
  assert_comment(comment)

  if (is.null(comment$style)) {
    rPr <- create_font()
  } else {
    rPr <- comment$style
  }

  rPr <- gsub("font>", "rPr>", rPr)
  sheet <- wb$.__enclos_env__$private$get_sheet_index(sheet)

  if (!is.null(dims)) {
    ref <- dims
    col <- col2int(dims_to_rowcol(dims)[["col"]])
    row <- as.integer(dims_to_rowcol(dims)[["row"]])
  } else {
    if (!is.numeric(col)) {
      col <- col2int(col)
    }
    ref <- paste0(int2col(col), row)
  }

  comment_list <- list(list(
    "ref" = ref,
    "author" = comment$author,
    "comment" = comment$text,
    "style" = rPr,
    "clientData" = genClientData(col, row, visible = comment$visible, height = comment$height, width = comment$width)
  ))

  # guard against integer(0) which is returned if no comment is found
  iterator <- function(x) {
    assert_class(x, "integer")
    if (length(x) == 0) x <- 0
    max(x) + 1
  }

  # check if relationships for this sheet already has comment entry and get next free rId
  if (length(wb$worksheets_rels[[sheet]]) == 0) wb$worksheets_rels[[sheet]] <- genBaseSheetRels(sheet)

  next_rid <- 1
  next_id  <- wb$worksheets[[sheet]]$relships$comments

  if (!all(identical(wb$worksheets_rels[[sheet]], character()))) {
    rels     <- rbindlist(xml_attr(wb$worksheets_rels[[sheet]], "Relationship"))
    rels$typ <- basename(rels$Type)
    rels$id  <- as.integer(gsub("\\D+", "", rels$Id))
    next_rid <- iterator(rels$id)
  }

  id <- 1025 + sum(lengths(wb$comments))

  fillcolor <- color %||% "#ffffe1"
  # looks like vml accepts only "#RGB" and not "ARGB"
  if (is_wbColour(fillcolor)) {
    if (names(fillcolor) != "rgb") {
      # actually there are more colors like: "lime [11]" and
      # "infoBackground [80]" (the default). But no clue how
      # these are created.
      stop("fillcolor needs to be an RGB color")
    }

    fillcolor <- paste0("#", substr(fillcolor, 3, 8))
  }

  rID <- NULL
  if (!is.null(file)) {
    vml_id <- wb$worksheets[[sheet]]$relships$vmlDrawing
    wb$.__enclos_env__$private$add_media(file = file)
    file <- names(wb$media)[length(wb$media)]
    if (length(vml_id)) # continue rId indexing
      rID <- paste0("rId", length(wb$vml_rels[[vml_id]]) + 1L)
    else # add a new rId
      rID <- paste0("rId", 1L)

    vml_relship <- sprintf(
      '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/%s"/>',
      rID,
      file
    )
  }


  # create new commment vml
  cd <- unapply(comment_list, "[[", "clientData")
  vml_xml <- read_xml(genBaseShapeVML(cd, id, fillcolor, rID), pointer = FALSE)
  vml_comment <- '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
  vml_xml <- paste0(vml_xml, vml_comment)

  # if this sheet has no comment entry in relationships, add a new relationship
  # 1) to Content_Types
  # 2) to worksheets_rels
  if (length(wb$worksheets[[sheet]]$relships$comments) == 0) {
    next_id <- length(wb$comments) + 1L

    wb$Content_Types <- c(
      wb$Content_Types,
      sprintf(
        '<Override PartName="/xl/comments%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>',
        next_id
      )
    )

    wb$worksheets[[sheet]]$relships$comments <- next_id

    # check if we have a vmlDrawing attached to the worksheet
    # if no) create one
    # if yes) update it
    if (length(wb$worksheets[[sheet]]$relships$vmlDrawing) == 0) {

      wb$worksheets_rels[[sheet]] <- c(
        wb$worksheets_rels[[sheet]],
        sprintf(
          '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing%s.vml"/>',
          next_rid,
          next_id
        )
      )

      ## create vml for output
      vml_xml <-  xml_node_create(
        xml_name = "xml",
        xml_attributes = c(
          `xmlns:v` = "urn:schemas-microsoft-com:vml",
          `xmlns:o` = "urn:schemas-microsoft-com:office:office",
          `xmlns:x` = "urn:schemas-microsoft-com:office:excel"
        ),
        xml_children = vml_xml
      )
      if (length(wb$vml) == 0) {
        wb$vml <- list()
      }
      wb$vml <- c(wb$vml, vml_xml)

      wb$worksheets[[sheet]]$relships$vmlDrawing <- next_id
      if (!is.null(rID)) {
        if (length(wb$vml_rels) == 0) {
          wb$vml_rels <- list()
        }
        if (length(wb$vml_rels) < next_id) {
          wb$vml_rels <- wb$vml_rels[seq_len(next_id)]
        }

        wb$vml_rels[[next_id]] <- append(
          wb$vml_rels[[next_id]],
          vml_relship
        )
      }
      # TODO hardcoded 2. Marvin fears that this is not good enough
      wb$worksheets[[sheet]]$legacyDrawing <- sprintf('<legacyDrawing r:id="rId%s"/>', next_rid)

      next_rid <- next_rid + 1
    } else {
      vml_id <- wb$worksheets[[sheet]]$relships$vmlDrawing
      wb$vml[[vml_id]] <- xml_add_child(wb$vml[[vml_id]], vml_xml)
      if (!is.null(rID)) {
        wb$vml_rels[[vml_id]] <- append(
          wb$vml_rels[[vml_id]],
          vml_relship
        )
      }
    }

    wb$worksheets_rels[[sheet]] <- c(
      wb$worksheets_rels[[sheet]],
      sprintf(
        '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments%s.xml"/>',
        next_rid,
        next_id
      )
    )
  } else {
    vml_id <- wb$worksheets[[sheet]]$relships$vmlDrawing
    wb$vml[[vml_id]] <- xml_add_child(wb$vml[[vml_id]], vml_xml)
      if (!is.null(rID)) {
        wb$vml_rels[[vml_id]] <- append(
          wb$vml_rels[[vml_id]],
          vml_relship
        )
      }
  }

  cmmnt_id <- wb$worksheets[[sheet]]$relships$comments

  if (length(wb$comments) == 0) {
    wb$comments <- list(NA)
  } else if (length(wb$comments) < cmmnt_id) {
    wb$comments <- append(wb$comments, NA)
  }

  if (all(is.na(wb$comments[[cmmnt_id]]))) {
    previous_comment <- NULL
  } else {
    previous_comment <- wb$comments[[cmmnt_id]]
  }

  wb$comments[[cmmnt_id]] <- c(previous_comment, comment_list)

  invisible(wb)
}

do_remove_comment <- function(
    wb,
    sheet,
    col        = NULL,
    row        = NULL,
    gridExpand = TRUE,
    dims       = NULL
) {
  # TODO add as method; wbWorkbook$remove_comment()
  assert_workbook(wb)

  sheet <- wb$.__enclos_env__$private$get_sheet_index(sheet)

  if (!is.null(col) && !is.null(row)) {
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
  }

  if (!is.null(dims)) {
    comb <- unlist(dims_to_dataframe(dims, fill = TRUE))
  }

  toKeep <- !sapply(wb$comments[[sheet]], "[[", "ref") %in% comb

  # FIXME: if all comments are removed we should drop to wb$comments <- list()
  wb$comments[[sheet]] <- wb$comments[[sheet]][toKeep]

}

as_fmt_txt <- function(x) {
  vapply(x, function(y) {
    ifelse(is_xml(y), si_to_txt(xml_node_create("si", xml_children = y)), y)
  },
  NA_character_
  )
}
