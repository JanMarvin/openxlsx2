wb_add_conditional_formatting_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    cols,
    rows,
    rule  = NULL,
    style = NULL,
    # TODO add vector of possible values
    type = c("expression", "colorScale", "dataBar", "duplicatedValues",
             "containsText", "notContainsText", "beginsWith", "endsWith",
             "between", "topN", "bottomN"),
    params = list(
      showValue = TRUE,
      gradient  = TRUE,
      border    = TRUE,
      percent   = FALSE,
      rank      = 5L
    )
) {
  if (!is.null(style)) assert_class(style, "character")
  assert_class(type, "character")
  assert_class(params, "list")

  type <- match.arg(type)

  ## rows and cols
  if (!is.numeric(cols)) {
    cols <- col2int(cols)
  }

  rows <- as.integer(rows)

  ## check valid rule
  dxfId <- NULL
  if (!is.null(style)) dxfId <- self$styles_mgr$get_dxf_id(style)
  params <- validate_cf_params(params)
  values <- NULL

  sel <- c("expression", "duplicatedValues", "containsText", "notContainsText", "beginsWith", "endsWith", "between", "topN", "bottomN")
  if (is.null(style) && type %in% sel) {
    smp <- random_string()
    style <- create_dxfs_style(font_color = wb_colour(hex = "FF9C0006"), bgFill = wb_colour(hex = "FFFFC7CE"))
    self$styles_mgr$add(style, smp)
    dxfId <- self$styles_mgr$get_dxf_id(smp)
  }

  switch(
    type,

    expression = {
      # TODO should we bother to do any conversions or require the text
      # entered to be exactly as an Excel expression would be written?
      msg <- "When type == 'expression', "

      if (!is.character(rule) || length(rule) != 1L) {
        stop(msg, "rule must be a single length character vector")
      }

      rule <- gsub("!=", "<>", rule)
      rule <- gsub("==", "=", rule)
      rule <- replace_legal_chars(rule) # replaces <>

      if (!grepl("[A-Z]", substr(rule, 1, 2))) {
        ## formula looks like "operatorX" , attach top left cell to rule
        rule <- paste0(
          get_cell_refs(data.frame(min(rows), min(cols))),
          rule
        )
      } ## else, there is a letter in the formula and apply as is

    },

    colorScale = {
      # - style is a vector of colours with length 2 or 3
      # - rule specifies the quantiles (numeric vector of length 2 or 3), if NULL min and max are used
      msg <- "When type == 'colourScale', "

      if (!is.character(style)) {
        stop(msg, "style must be a vector of colours of length 2 or 3.")
      }

      if (!length(style) %in% 2:3) {
        stop(msg, "style must be a vector of length 2 or 3.")
      }

      if (!is.null(rule)) {
        if (length(rule) != length(style)) {
          stop(msg, "rule and style must have equal lengths.")
        }
      }

      style <- check_valid_colour(style)

      if (isFALSE(style)) {
        stop(msg, "style must be valid colors")
      }

      values <- rule
      rule <- style
    },

    dataBar = {
      # - style is a vector of colours of length 2 or 3
      # - rule specifies the quantiles (numeric vector of length 2 or 3), if NULL min and max are used
      msg <- "When type == 'dataBar', "
      style <- style %||% "#638EC6"

      # TODO use inherits() not class()
      if (!inherits(style, "character")) {
        stop(msg, "style must be a vector of colours of length 1 or 2.")
      }

      if (!length(style) %in% 1:2) {
        stop(msg, "style must be a vector of length 1 or 2.")
      }

      if (!is.null(rule)) {
        if (length(rule) != length(style)) {
          stop(msg, "rule and style must have equal lengths.")
        }
      }

      ## Additional parameters passed by ...
      # showValue, gradient, border
      style <- check_valid_colour(style)

      if (isFALSE(style)) {
        stop(msg, "style must be valid colors")
      }

      values <- rule
      rule <- style
    },

    duplicatedValues = {
      # type == "duplicatedValues"
      # - style is a Style object
      # - rule is ignored

      rule <- style
    },

    containsText = {
      # - style is Style object
      # - rule is text to look for
      msg <- "When type == 'contains', "

      if (!inherits(rule, "character")) {
        stop(msg, "rule must be a character vector of length 1.")
      }

      values <- rule
      rule <- style
    },

    notContainsText = {
      # - style is Style object
      # - rule is text to look for
      msg <- "When type == 'notContains', "

      if (!inherits(rule, "character")) {
        stop(msg, "rule must be a character vector of length 1.")
      }

      values <- rule
      rule <- style
    },

    beginsWith = {
      # - style is Style object
      # - rule is text to look for
      msg <- "When type == 'beginsWith', "

      if (!is.character("character")) {
        stop(msg, "rule must be a character vector of length 1.")
      }

      values <- rule
      rule <- style
    },

    endsWith = {
      # - style is Style object
      # - rule is text to look for
      msg <- "When type == 'endsWith', "

      if (!inherits(rule, "character")) {
        stop(msg, "rule must be a character vector of length 1.")
      }

      values <- rule
      rule <- style
    },

    between = {
      rule <- range(rule)
    },

    topN = {
      # - rule is ignored
      # - 'rank' and 'percent' are named params

      ## Additional parameters passed by ...
      # percent, rank

      values <- params
      rule <- style
    },

    bottomN = {
      # - rule is ignored
      # - 'rank' and 'percent' are named params

      ## Additional parameters passed by ...
      # percent, rank

      values <- params
      rule <- style
    }
  )

  private$do_conditional_formatting(
    sheet    = sheet,
    startRow = min(rows),
    endRow   = max(rows),
    startCol = min(cols),
    endCol   = max(cols),
    dxfId    = dxfId,
    formula  = rule,
    type     = type,
    values   = values,
    params   = params
  )

  invisible(self)
}
