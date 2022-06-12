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
#' @param ... See below
#' @details See Examples.
#'
#' If type == "expression"
#' \itemize{
#'   \item{style is a Style object.}
#'   \item{rule is an expression. Valid operators are "<", "<=", ">", ">=", "==", "!=".}
#' }
#'
#' If type == "colourScale"
#' \itemize{
#'   \item{style is a vector of colours with length 2 or 3}
#'   \item{rule can be NULL or a vector of colours of equal length to styles}
#' }
#'
#' If type == "databar"
#' \itemize{
#'   \item{style is a vector of colours with length 2 or 3}
#'   \item{rule is a numeric vector specifying the range of the databar colours. Must be equal length to style}
#'   \item{...
#'   \itemize{
#'     \item{**showvalue** If FALSE the cell value is hidden. Default TRUE.}
#'     \item{**gradient** If FALSE colour gradient is removed. Default TRUE.}
#'     \item{**border** If FALSE the border around the database is hidden. Default TRUE.}
#'      }
#'    }
#' }
#'
#' If type == "duplicates"
#' \itemize{
#'   \item{style is a Style object.}
#'   \item{rule is ignored.}
#' }
#'
#' If type == "contains"
#' \itemize{
#'   \item{style is a Style object.}
#'   \item{rule is the text to look for within cells}
#' }
#'
#' If type == "between"
#' \itemize{
#'   \item{style is a Style object.}
#'   \item{rule is a numeric vector of length 2 specifying lower and upper bound (Inclusive)}
#' }
#'
#' If type == "topN"
#' \itemize{
#'   \item{style is a Style object.}
#'   \item{rule is ignored}
#'   \item{...
#'   \itemize{
#'     \item{**rank** numeric vector of length 1 indicating number of highest values.}
#'     \item{**percent** TRUE if you want top N percentage.}
#'      }
#'    }
#' }
#'
#' If type == "bottomN"
#' \itemize{
#'   \item{style is a Style object.}
#'   \item{rule is ignored}
#'   \item{...
#'   \itemize{
#'     \item{**rank** numeric vector of length 1 indicating number of lowest values.}
#'     \item{**percent** TRUE if you want bottom N percentage.}
#'      }
#'    }
#' }
#'
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("a")
#' wb$add_conditional_formatting("a", 1, 1, )
#'
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
    ...
) {
    assert_workbook(wb)
    wb$add_conditional_formatting(
      wb    = wb,
      sheet = sheet,
      cols  = cols,
      rows  = rows,
      rule  = rule,
      style = style,
      type  = type,
      ...
    )
  }

#' @rdname wb_add_conditional_formatting
#' @export
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
  wb_add_conditional_formatting(
    wb    = wb,
    sheet = sheet,
    cols  = cols,
    rows  = rows,
    rule  = rule,
    style = style,
    type  = type,
    ...
  )
}
