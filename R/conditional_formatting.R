#' @name conditionalFormatting
#' @aliases databar
#' @title Add conditional formatting to cells
#' @description Add conditional formatting to cells
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Columns to apply conditional formatting to
#' @param rows Rows to apply conditional formatting to
#' @param rule The condition under which to apply the formatting. See examples.
#' @param style A style to apply to those cells that satisfy the rule. Default is 'font_color = "FF9C0006"' and 'bgFill = "FFFFC7CE"'
#' @param type Either 'expression', 'colourScale', 'databar', 'duplicates', 'beginsWith',
#' 'endsWith', 'topN', 'bottomN', 'contains' or 'notContains' (case insensitive).
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
#' wb$add_worksheet("cellIs")
#'
#' negStyle <- create_dxfs_style(font_color = c(rgb = "FF9C0006"), bgFill = c(rgb = "FFFFC7CE"))
#' posStyle <- create_dxfs_style(font_color = c(rgb = "FF006100"), bgFill = c(rgb = "FFC6EFCE"))
#'
#' wb$styles_mgr$styles$dxfs <- c(wb$styles_mgr$styles$dxfs,
#'                                 c(negStyle, posStyle)
#'                              )
#'
#' set.seed(123)
#'
#' ## rule applies to all each cell in range
#' wb$add_data("cellIs", -5:5)
#' wb$add_data("cellIs", LETTERS[1:11], startCol = 2)
#' wb_conditional_formatting(wb, "cellIs",
#'                       cols = 1,
#'                       rows = 1:11, rule = "!=0", style = negStyle
#' )
#'
#' wb_conditional_formatting(wb, "cellIs",
#'                       cols = 1,
#'                       rows = 1:11, rule = "==0", style = posStyle
#' )
#'
#'
#' wb$add_worksheet("Moving Row")
#' ## highlight row dependent on first cell in row
#' wb$add_data("Moving Row", -5:5)
#' wb$add_data("Moving Row", LETTERS[1:11], startCol = 2)
#' wb_conditional_formatting(wb, "Moving Row",
#'                       cols = 1:2,
#'                       rows = 1:11, rule = "$A1<0", style = negStyle
#' )
#' wb_conditional_formatting(wb, "Moving Row",
#'                       cols = 1:2,
#'                       rows = 1:11, rule = "$A1>0", style = posStyle
#' )
#'
#'
#' wb$add_worksheet("Moving Col")
#' ## highlight column dependent on first cell in column
#' wb$add_data("Moving Col", -5:5)
#' wb$add_data("Moving Col", LETTERS[1:11], startCol = 2)
#' wb_conditional_formatting(wb, "Moving Col",
#'                       cols = 1:2,
#'                       rows = 1:11, rule = "A$1<0", style = negStyle
#' )
#' wb_conditional_formatting(wb, "Moving Col",
#'                       cols = 1:2,
#'                       rows = 1:11, rule = "A$1>0", style = posStyle
#' )
#'
#'
#' wb$add_worksheet("Dependent on")
#' ## highlight entire range cols X rows dependent only on cell A1
#' wb$add_data("Dependent on", -5:5)
#' wb$add_data("Dependent on", LETTERS[1:11], startCol = 2)
#' wb_conditional_formatting(wb, "Dependent on",
#'                       cols = 1:2,
#'                       rows = 1:11, rule = "$A$1<0", style = negStyle
#' )
#' wb_conditional_formatting(wb, "Dependent on",
#'                       cols = 1:2,
#'                       rows = 1:11, rule = "$A$1>0", style = posStyle
#' )
#'
#'
#' ## highlight cells in column 1 based on value in column 2
#' wb$add_data("Dependent on", data.frame(x = 1:10, y = runif(10)), startRow = 15)
#' wb_conditional_formatting(wb, "Dependent on",
#'                       cols = 1,
#'                       rows = 16:25, rule = "B16<0.5", style = negStyle
#' )
#' wb_conditional_formatting(wb, "Dependent on",
#'                       cols = 1,
#'                       rows = 16:25, rule = "B16>=0.5", style = posStyle
#' )
#'
#'
#' wb$add_worksheet("Duplicates")
#' ## highlight duplicates using default style
#' wb$add_data("Duplicates", sample(LETTERS[1:15], size = 10, replace = TRUE))
#' wb_conditional_formatting(wb, "Duplicates", cols = 1, rows = 1:10, type = "duplicates")
#'
#'
#' wb$add_worksheet("containsText")
#' ## cells containing text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' wb$add_data("containsText", sapply(1:10, fn))
#' wb_conditional_formatting(wb, "containsText", cols = 1, rows = 1:10, type = "contains", rule = "A")
#'
#'
#' wb$add_worksheet("notcontainsText")
#' ## cells not containing text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' wb$add_data("containsText", sapply(1:10, fn))
#' wb_conditional_formatting(wb, "notcontainsText", cols = 1,
#'                       rows = 1:10, type = "notcontains", rule = "A")
#'
#'
#' wb$add_worksheet("beginsWith")
#' ## cells begins with text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' wb$add_data("beginsWith", sapply(1:100, fn))
#' wb_conditional_formatting(wb, "beginsWith", cols = 1, rows = 1:100, type = "beginsWith", rule = "A")
#'
#'
#' wb$add_worksheet("endsWith")
#' ## cells ends with text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' wb$add_data("endsWith", sapply(1:100, fn))
#' wb_conditional_formatting(wb, "endsWith", cols = 1, rows = 1:100, type = "endsWith", rule = "A")
#'
#' wb$add_worksheet("colourScale", zoom = 30)
#' ## colourscale colours cells based on cell value
#' df <- read_xlsx(system.file("extdata", "readTest.xlsx", package = "openxlsx2"), sheet = 4)
#' wb$add_data("colourScale", df, colNames = FALSE) ## write data.frame
#' ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
#' ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
#' wb_conditional_formatting(wb, "colourScale",
#'                       cols = seq_along(df), rows = 1:nrow(df),
#'                       style = c("black", "white"),
#'                       rule = c(0, 255),
#'                       type = "colourScale"
#' )
#' wb_set_col_widths(wb, "colourScale", cols = seq_along(df), widths = 1.07)
#' wb <- wb_set_row_heights(wb, "colourScale", rows = seq_len(nrow(df)), heights = 7.5)
#'
#' wb$add_worksheet("databar")
#' ## Databars
#' wb$add_data("databar", -5:5)
#' wb_conditional_formatting(wb, "databar", cols = 1, rows = 1:11, type = "databar") ## Default colours
#'
#' wb$add_worksheet("between")
#' ## Between
#' # Highlight cells in interval [-2, 2]
#' wb$add_data("between", -5:5)
#' wb_conditional_formatting(wb, "between", cols = 1, rows = 1:11, type = "between", rule = c(-2, 2))
#'
#' wb$add_worksheet("topN")
#' ## Top N
#' wb$add_data("topN", data.frame(x = 1:10, y = rnorm(10)))
#' # Highlight top 5 values in column x
#' wb_conditional_formatting(wb, "topN", cols = 1, rows = 2:11,
#'                       style = posStyle, type = "topN", rank = 5)#'
#' # Highlight top 20 percentage in column y
#' wb_conditional_formatting(wb, "topN", cols = 2, rows = 2:11,
#'                       style = posStyle, type = "topN", rank = 20, percent = TRUE)
#'
#' wb$add_worksheet("bottomN")
#' ## Bottom N
#' wb$add_data("bottomN", data.frame(x = 1:10, y = rnorm(10)))
#' # Highlight bottom 5 values in column x
#' wb_conditional_formatting(wb, "bottomN", cols = 1, rows = 2:11,
#'                       style = negStyle, type = "topN", rank = 5)
#' # Highlight bottom 20 percentage in column y
#' wb_conditional_formatting(wb, "bottomN", cols = 2, rows = 2:11,
#'                       style = negStyle, type = "topN", rank = 20, percent = TRUE)
#'
#' wb$add_worksheet("logical operators")
#' ## Logical Operators
#' # You can use Excels logical Operators
#' wb$add_data("logical operators", 1:10)
#' wb_conditional_formatting(wb, "logical operators",
#'                       cols = 1, rows = 1:10,
#'                       rule = "OR($A1=1,$A1=3,$A1=5,$A1=7)"
#' )
#'
wb_conditional_formatting <- function(
    wb,
    sheet,
    cols,
    rows,
    rule = NULL,
    style = NULL,
    type = "expression",
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
