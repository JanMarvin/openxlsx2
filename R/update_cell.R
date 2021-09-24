
#' replace data in a single cell
#'
#' Minimal invasive update of a single cell for imported workbooks.
#'
#' @param x value you want to insert
#' @param wb the workbook you want to update
#' @param sheet the sheet you want to update
#' @param cell the cell you want to update in Excel conotation e.g. "A1"
#' @param data_class optional data class object
#' @param colNames if TRUE colNames are passed down
#'
#' @examples
#'    xlsxFile <- system.file("extdata", "update_test.xlsx", package = "openxlsx2")
#'    wb <- loadWorkbook(xlsxFile)
#'
#'    # update Cells D4:D6 with 1:3
#'    wb <- update_cell(x = c(1:3), wb = wb, sheet = "Sheet1", cell = "D4:D6")
#'
#'    # update Cells B3:D3 (names())
#'    wb <- update_cell(x = c("x", "y", "z"), wb = wb, sheet = "Sheet1", cell = "B3:D3")
#'
#'    # update D4 again (single value this time)
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "D4")
#'
#'    # add new column on the left of the existing workbook
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "A4")
#'
#'    # add new row on the end of the existing workbook
#'    wb <- update_cell(x = 7, wb = wb, sheet = "Sheet1", cell = "A9")
#'    wb_to_df(wb)
#'
#' @export
update_cell <- function(x, wb, sheet, cell, data_class, colNames = FALSE) {

  dimensions <- unlist(strsplit(cell, ":"))
  rows <- gsub("[[:upper:]]","", dimensions)
  cols <- gsub("[[:digit:]]","", dimensions)

  if (length(dimensions) == 2) {
    # cols
    cols <- col2int(cols)
    cols <- seq(cols[1], cols[2])
    cols <- int2col(cols)

    rows <- as.character(seq(rows[1], rows[2]))
  }


  if(is.character(sheet)) {
    sheet_id <- which(sheet == wb$sheet_names)
  } else {
    sheet_id <- sheet
  }

  if (missing(data_class))
    data_class <- sapply(x, class)

  # if(identical(sheet_id, integer(0)))
  #   stop("sheet not in workbook")

  # 1) pull sheet to modify from workbook; 2) modify it; 3) push it back
  cc  <- wb$worksheets[[sheet_id]]$sheet_data$cc
  row_attr <- wb$worksheets[[sheet_id]]$sheet_data$row_attr

  # workbooks contain only entries for values currently present.
  # if A1 is filled, B1 is not filled and C1 is filled the sheet will only
  # contain fields A1 and C1.
  cc$row <- paste0(cc$c_r, cc$row_r)
  cells_in_wb <- cc$rw
  rows_in_wb <- names(row_attr)


  # check if there are rows not available
  if (!all(rows %in% rows_in_wb)) {
    # message("row(s) not in workbook")

    # add row to name vector, extend the entire thing
    total_rows <- as.character(sort(unique(as.numeric(c(rows, rows_in_wb)))))

    # new row_attr
    row_attr_new <- vector("list", length(rows_in_wb))
    names(row_attr_new) <- rows_in_wb

    row_attr_new[rows_in_wb] <- row_attr[rows_in_wb]

    for (trow in total_rows) {
      row_attr_new[[trow]] <- list(r = trow)
    }

    wb$worksheets[[sheet_id]]$sheet_data$row_attr <- row_attr_new
    # provide output
    rows_in_wb <- total_rows

  }

  if (!any(cols %in% cells_in_wb)) {
    # all rows are availabe in the dataframe
    for (row in rows){

      # collect all wanted cols and order for excel
      total_cols <- unique(c(cc$c_r[cc$row_r == row], cols))
      total_cols <- int2col(sort(col2int(total_cols)))

      # create candidate
      cc_row_new <- data.frame(matrix(NA_character_, nrow = length(total_cols), ncol = 2))
      names(cc_row_new) <- names(cc)[1:2]
      cc_row_new$row_r <- row
      cc_row_new$c_r <- total_cols

      # extract row (easier or maybe only way to change order?)
      cc_row <- cc[cc$row_r == row, ]
      # remove row from cc
      if (nrow(cc_row)>0) cc <- cc[-which(rownames(cc) %in% rownames(cc_row)),]
      # new row
      cc_row <- merge(x = cc_row_new, y = cc_row, all.x = TRUE)

      # assign to cc
      cc <- rbind(cc, cc_row)

    }
  }

  # update dimensions
  cc$row <- paste0(cc$c_r, cc$row_r)
  cells_in_wb <- cc$rw

  all_rows <- unique(cc$row_r)
  all_cols <- unique(cc$c_r)

  min_cell <- trimws(paste0(min(all_cols), min(all_rows)))
  max_cell <- trimws(paste0(max(all_cols), max(all_rows)))

  # i know, i know, i'm lazy
  wb$worksheets[[sheet_id]]$dimension <- paste0("<dimension ref=\"", min_cell, ":", max_cell, "\"/>")

  # if (any(rows %in% rows_in_wb) )
    # message("found cell(s) to update")

  if (all(rows %in% rows_in_wb)) {
    # message("cell(s) to update already in workbook. updating ...")

    i <- 0; n <- 0
    for (row in rows) {

      n <- n+1
      m <- 0

      for (col in cols) {
        i <- i+1
        m <- m+1

        # check if is data frame or matrix
        value <- ifelse(is.null(dim(x)), x[i], x[n, m])

        sel <- cc$row_r == row & cc$c_r == col
        cc[sel, c("c_s", "c_t", "v", "f", "f_t", "t")] <- "_openxlsx_NA_"


        # for now convert all R-characters to inlineStr (e.g. names() of a dataframe)
        if (data_class[m] %in% c("character", "factor") | (colNames == TRUE & n == 1)) {
          cc[sel, "c_t"] <- "inlineStr"
          cc[sel, "t"]   <- as.character(value)
        } else {
          cc[sel, "v"]   <- as.character(value)
        }

      }
    }

  }

  # order cc
  cc$ordered_rows <- col2int(cc$c_r)
  cc$ordered_cols <- as.numeric(cc$row_r)

  cc <- cc[order(cc$ordered_rows, cc$ordered_cols),]
  cc$ordered_cols <- NULL
  cc$ordered_rows <- NULL

  # push everything back to workbook
  wb$worksheets[[sheet_id]]$sheet_data$cc  <- cc

  wb
}
