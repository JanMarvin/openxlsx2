#' Add conditional formatting to cells
#'
#' Add conditional formatting to cells

#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Columns to apply conditional formatting to
#' @param rows Rows to apply conditional formatting to
#' @param rule The condition under which to apply the formatting. See examples.
#' @param style A style to apply to those cells that satisfy the rule. Default is 'font_color = "FF9C0006"' and 'bgFill = "FFFFC7CE"'
#' @param type The type of conditional formatting rule to apply.
#' @param params Additional parameters passed.  See **Details** for more
#' @details See Examples.
#'
#' @details
#' Conditional formatting types accept different parameters.  Unless noted,
#' unlisted parameters are ignored.
#'
#' \describe{
#'   \item{`expression`}{
#'     `[style]`\cr A `Style` object\cr\cr
#'     `[rule]`\cr An Excel expression (as a character). Valid operators are: `<`, `<=`, `>`, `>=`, `==`, `!=`
#'   }
#'   \item{colorScale}{
#'     `[style]`\cr A `character` vector of valid colors with length `2` or `3`\cr\cr
#'     `[rule]`\cr `NULL` or a `character` vector of valid colors of equal length to `styles`
#'   }
#'   \item{dataBar}{
#'     `[style]`\cr A `character` vector of valid colors with length `2` or `3`\cr\cr
#'     `[rule]`\cr A `numeric` vector specifying the range of the databar colors. Must be equal length to `style`\cr\cr
#'     `[params$showValue]`\cr If `FALSE` the cell value is hidden. Default `TRUE`\cr\cr
#'     `[params$gradient]`\cr If `FALSE` colour gradient is removed. Default `TRUE`\cr\cr
#'     `[params$border]`\cr If `FALSE` the border around the database is hidden. Default `TRUE`
#'   }
#'   \item{duplicated}{
#'     `[style]`\cr A `Style` object
#'   }
#'   \item{contains}{
#'     `[style]`\cr A `Style` object\cr\cr
#'     `[rule]`\cr The text to look for within cells
#'   }
#'   \item{between}{
#'     `[style]`\cr A Style object.\cr\cr
#'     `[rule]`\cr A `numeric` vector of length `2` specifying lower and upper bound (Inclusive)
#'   }
#'   \item{topN}{
#'     `[style]`\cr A `Style` object\cr\cr
#'     `[params$rank]`\cr A `numeric` vector of length `1` indicating number of highest values\cr\cr
#'     `[params$percent]` If `TRUE` uses percentage
#'   }
#'   \item{bottomN}{
#'     `[style]`\cr A `Style` object\cr\cr
#'     `[params$rank]`\cr A `numeric` vector of length `1` indicating number of lowest values\cr\cr
#'     `[params$percent]`\cr If `TRUE` uses percentage
#'   }
#' }
#'
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("a")
#' wb$add_conditional_formatting("a", 1, 1, )
wb_add_conditional_formatting <- function(
    wb,
    sheet,
    cols,
    rows,
    rule = NULL,
    style = NULL,
    type = c("expression", "colorScale", "dataBar", "duplicatedValues",
             "containsText", "notContainsText", "beginsWith", "endsWith",
             "between", "topN", "bottomN"),
    params = list(
      showValue = TRUE,
      gradient  = TRUE,
      border    = TRUE,
      percent   = FALSE,
      rank      = NULL
    )
) {
    assert_workbook(wb)
    wb$add_conditional_formatting(
      sheet = sheet,
      cols  = cols,
      rows  = rows,
      rule  = rule,
      style = style,
      type  = type,
      params = params
    )
  }

#' @rdname wb_add_conditional_formatting
#' @export
#' @param ... passed to `params`
wb_conditional_formatting <- function(
    wb,
    sheet,
    cols,
    rows,
    rule = NULL,
    style = NULL,
    type = c("expression", "colorScale", "dataBar", "duplicatedValues",
             "containsText", "notContainsText", "beginsWith", "endsWith",
             "between", "topN", "bottomN"),
    ...
) {
  .Deprecated("wb_add_conditional_formatting()")

  params <- list(...)
  params$showValue <- params$showValue %||% TRUE
  params$gradient  <- params$gradient  %||% TRUE
  params$border    <- params$border    %||% TRUE
  params$percent   <- params$percent   %||% FALSE
  params$percent   <- params$percent   %||% NULL

  wb_add_conditional_formatting(
    wb     = wb,
    sheet  = sheet,
    cols   = cols,
    rows   = rows,
    rule   = rule,
    style  = style,
    type   = type,
    params = params
  )
}
