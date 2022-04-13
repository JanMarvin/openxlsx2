
#' Create a new Workbook object
#'
#' Create a new Workbook object
#'
#' @param creator Creator of the workbook (your name). Defaults to login username
#' @param title Workbook properties title
#' @param subject Workbook properties subject
#' @param category Workbook properties category
#' @param datetimeCreated The time of the workbook is created
#' @return A [wbWorkbook] object
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Save workbook to working directory
#' \dontrun{
#' wb_save(wb, path = temp_xlsx(), overwrite = TRUE)
#' }
#'
#' ## Set Workbook properties
#' wb <- wb_workbook(
#'   creator  = "Me",
#'   title    = "Expense Report",
#'   subject  = "Expense Report - 2022 Q1",
#'   category = "sales"
#' )
wb_workbook <- function(
  creator         = NULL,
  title           = NULL,
  subject         = NULL,
  category        = NULL,
  datetimeCreated = Sys.time()
) {
  wbWorkbook$new(
    creator         = creator,
    title           = title,
    subject         = subject,
    category        = category,
    datetimeCreated = datetimeCreated
  )
}


#' Save Workbook to file
#'
#' @param wb A `wbWorkbook` object to write to file
#' @param path A path to save the workbook to
#' @param overwrite If `FALSE`, will not overwrite when `path` exists
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' ## Create a new workbook and add a worksheet
#' wb <- wb_workbook("Creator of workbook")
#' wb$addWorksheet(sheet = "My first worksheet")
#'
#' ## Save workbook to working directory
#' \dontrun{
#' wb_save(wb, path = temp_xlsx(), overwrite = TRUE)
#' }
wb_save <- function(wb, path = NULL, overwrite = TRUE) {
  assert_workbook(wb)
  wb$clone()$save(path = path, overwrite = overwrite)$path
}


#' Worksheet cell merging
#'
#' Merge cells within a worksheet
#'
#' @details As merged region must be rectangular, only min and max of cols and
#'   rows are used.
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols,rows Column and row specifications to merge on.  Note: `min()` and
#'   `max()` of each vector are provided for specs.  Skipping rows or columns is
#'   not recognized.
#'
#' @examples
#' # Create a new workbook
#' wb <- wb_workbook()
#'
#' # Add a worksheets
#' wb$addWorksheet("Sheet 1")
#' wb$addWorksheet("Sheet 2")
#'
#' # Merge cells: Row 2 column C to F (3:6)
#' wb <- wb_merge_cells(wb, "Sheet 1", cols = 2, rows = 3:6)
#'
#' # Merge cells:Rows 10 to 20 columns A to J (1:10)
#' wb <- wb_merge_cells(wb, 1, cols = 1:10, rows = 10:20)
#'
#' # Intersecting merges
#' wb <- wb_merge_cells(wb, 2, cols = 1:10, rows = 1)
#' wb <- wb_merge_cells(wb, 2, cols = 5:10, rows = 2)
#' wb <- wb_merge_cells(wb, 2, cols = c(1, 10), rows = 12) # equivalent to 1:10 as only min/max are used
#' try(wb_merge_cells(wb, 2, cols = 1, rows = c(1,10)))    # intersects existing merge
#'
#' # remove merged cells
#' wb <- wb_unmerge_cells(wb, 2, cols = 1, rows = 1)  # removes any intersecting merges
#' wb <- wb_merge_cells(wb, 2, cols = 1, rows = 1:10) # Now this works
#'
#' # Save workbook
#' \dontrun{
#' wb_save(wb, temp_xlsx(), overwrite = TRUE)
#' }
#'
#' @name ws_cell_merge
#' @family workbook wrappers
NULL

# merge cells -------------------------------------------------------------

#' @export
#' @rdname ws_cell_merge
wb_merge_cells <- function(wb, sheet, rows = NULL, cols = NULL) {
  assert_workbook(wb)
  wb$clone()$addCellMerge(sheet, rows = rows, cols = cols)
}

#' @export
#' @rdname ws_cell_merge
wb_unmerge_cells <- function(wb, sheet, rows = NULL, cols = NULL) {
  assert_workbook(wb)
  wb$clone()$removeCellMerge(sheet, rows = rows, cols = cols)
}


# worksheets --------------------------------------------------------------

#' Add a worksheet to a workbook
#'
#' @param wb A Workbook object to attach the new worksheet
#' @param sheet A name for the new worksheet
#' @param gridLines A logical. If `FALSE`, the worksheet grid lines will be
#'   hidden.
#' @param tabColour Colour of the worksheet tab. A valid colour (belonging to
#'   colours()) or a valid hex colour beginning with "#"
#' @param zoom A numeric between 10 and 400. Worksheet zoom level as a
#'   percentage.
#' @param header,oddHeader,evenHeader,firstHeader,footer,oddFooter,evenFooter,firstFooter
#'   Character vector of length 3 corresponding to positions left, center,
#'   right.  `header` and `footer` are used to default additional arguments.
#'   Setting `even`, `odd`, or `first`, overrides `header`/`footer`. Use `NA` to
#'   skip a position.
#' @param visible If FALSE, sheet is hidden else visible.
#' @param hasDrawing If TRUE prepare a drawing output (TODO does this work?)
#' @param paperSize An integer corresponding to a paper size. See ?pageSetup for
#'   details.
#' @param orientation One of "portrait" or "landscape"
#' @param hdpi Horizontal DPI. Can be set with options("openxlsx.dpi" = X) or
#'   options("openxlsx.hdpi" = X)
#' @param vdpi Vertical DPI. Can be set with options("openxlsx.dpi" = X) or
#'   options("openxlsx.vdpi" = X)
#' @details Headers and footers can contain special tags \itemize{
#'   \item{**&\[Page\]**}{ Page number} \item{**&\[Pages\]**}{ Number of pages}
#'   \item{**&\[Date\]**}{ Current date} \item{**&\[Time\]**}{ Current time}
#'   \item{**&\[Path\]**}{ File path} \item{**&\[File\]**}{ File name}
#'   \item{**&\[Tab\]**}{ Worksheet name} }
#' @return The [wbWorkbook] object `wb`
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook("Fred")
#'
#' ## Add 3 worksheets
#' wb$addWorksheet("Sheet 1")
#' wb$addWorksheet("Sheet 2", showGridLines = FALSE)
#' wb$addWorksheet("Sheet 3", tabColour = "red")
#' wb$addWorksheet("Sheet 4", showGridLines = FALSE, tabColour = "#4F81BD")
#'
#' ## Headers and Footers
#' wb$addWorksheet("Sheet 5",
#'   header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
#'   footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
#'   evenHeader = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
#'   evenFooter = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
#'   firstHeader = c("TOP", "OF FIRST", "PAGE"),
#'   firstFooter = c("BOTTOM", "OF FIRST", "PAGE")
#' )
#'
#' wb$addWorksheet("Sheet 6",
#'   header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
#'   footer = c("&[Path]&[File]", NA, "&[Tab]"),
#'   firstHeader = c(NA, "Center Header of First Page", NA),
#'   firstFooter = c(NA, "Center Footer of First Page", NA)
#' )
#'
#' wb$addWorksheet("Sheet 7",
#'   header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
#'   footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2")
#' )
#'
#' wb$addWorksheet("Sheet 8",
#'   firstHeader = c("FIRST ONLY L", NA, "FIRST ONLY R"),
#'   firstFooter = c("FIRST ONLY L", NA, "FIRST ONLY R")
#' )
#'
#' ## Need data on worksheet to see all headers and footers
#' writeData(wb, sheet = 5, 1:400)
#' writeData(wb, sheet = 6, 1:400)
#' writeData(wb, sheet = 7, 1:400)
#' writeData(wb, sheet = 8, 1:400)
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "addWorksheetExample.xlsx", overwrite = TRUE)
#' }
wb_add_worksheet <- function(
  wb,
  sheet,
  showGridLines = TRUE,
  tabColour     = NULL,
  zoom          = 100,
  header        = NULL,
  footer        = NULL,
  oddHeader     = header,
  oddFooter     = footer,
  evenHeader    = header,
  evenFooter    = footer,
  firstHeader   = header,
  firstFooter   = footer,
  visible       = c("true", "false", "hidden", "visible", "veryhidden"),
  hasDrawing    = FALSE,
  paperSize     = getOption("openxlsx.paperSize", default = 9),
  orientation   = getOption("openxlsx.orientation", default = "portrait"),
  hdpi          = getOption("openxlsx.hdpi", default = getOption("openxlsx.dpi", default = 300)),
  vdpi          = getOption("openxlsx.vdpi", default = getOption("openxlsx.dpi", default = 300))
) {
  assert_workbook(wb)
  wb$clone()$addWorksheet(
    sheet         = sheet,
    showGridLines = showGridLines,
    tabColour     = tabColour,
    zoom          = zoom,
    oddHeader     = headerFooterSub(oddHeader),
    oddFooter     = headerFooterSub(oddFooter),
    evenHeader    = headerFooterSub(evenHeader),
    evenFooter    = headerFooterSub(evenFooter),
    firstHeader   = headerFooterSub(firstHeader),
    firstFooter   = headerFooterSub(firstFooter),
    visible       = visible,
    paperSize     = paperSize,
    orientation   = orientation,
    vdpi          = vdpi,
    hdpi          = hdpi
  )
}


#' Clone a worksheet to a workbook
#'
#' Clone a worksheet to a Workbook object
#'
#' @param wb A [wbWorkbook] object
#' @param old Name of existing worksheet to copy
#' @param new Name of New worksheet to create
#' @return The `wb` object
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' # Create a new workbook
#' wb <- wb_workbook("Fred")
#'
#' # Add worksheets
#' wb$addWorksheet("Sheet 1")
#' wb$cloneWorksheet("Sheet 1", "Sheet 2")
#'
#' # Save workbook
#' \dontrun{
#' wb_save(wb, "cloneWorksheetExample.xlsx", overwrite = TRUE)
#' }
wb_clone_worksheet <- function(wb, old, new) {
  assert_workbook(wb)
  wb$clone()$cloneWorksheet(old = old, new = new)
}


#' @name freezePane
#' @title Freeze a worksheet pane
#' @description Freeze a worksheet pane
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param firstActiveRow Top row of active region
#' @param firstActiveCol Furthest left column of active region
#' @param firstRow If `TRUE`, freezes the first row (equivalent to firstActiveRow = 2)
#' @param firstCol If `TRUE`, freezes the first column (equivalent to firstActiveCol = 2)
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook("Kenshin")
#'
#' ## Add some worksheets
#' wb$addWorksheet("Sheet 1")
#' wb$addWorksheet("Sheet 2")
#' wb$addWorksheet("Sheet 3")
#' wb$addWorksheet("Sheet 4")
#'
#' ## Freeze Panes
#' wb$freezePanes("Sheet 1", firstActiveRow = 5, firstActiveCol = 3)
#' wb$freezePanes("Sheet 2", firstCol = TRUE) ## shortcut to firstActiveCol = 2
#' wb$freezePanes(3, firstRow = TRUE) ## shortcut to firstActiveRow = 2
#' wb$freezePanes(4, firstActiveRow = 1, firstActiveCol = "D")
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "freezePaneExample.xlsx", overwrite = TRUE)
#' }
wb_freeze_pane <- function(wb, sheet, firstActiveRow = NULL, firstActiveCol = NULL, firstRow = FALSE, firstCol = FALSE) {
  assert_workbook(wb)
  wb$clone()$freezePanes(
    sheet,
    firstActiveRow = firstActiveRow,
    firstActiveCol = firstActiveCol,
    firstRow       = firstRow,
    firstCol       = firstCol
  )
}


# heights and columns -----------------------------------------------------

# TODO order these...

#' Set worksheet row heights
#'
#' Set worksheet row heights
#'
#' @param wb A [wbWorkbook] object
#' @param sheet A name or index of a worksheet
#' @param rows Indices of rows to set height
#' @param heights Heights to set rows to specified in Excel column height units.
#'
#' @export
#' @family workbook wrappers
#' @seealso [removeRowHeights()]
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$addWorksheet("Sheet 1")
#'
#' ## set row heights
#' wb <- wb_set_row_heights(
#'   wb, 1,
#'   rows = c(1, 4, 22, 2, 19),
#'   heights = c(24, 28, 32, 42, 33)
#' )
#'
#' ## overwrite row 1 height
#' wb <- wb_set_row_heights(wb, 1, rows = 1, heights = 40)
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "setRowHeightsExample.xlsx", overwrite = TRUE)
#' }
wb_set_row_heights <- function(wb, sheet, rows, heights) {
  assert_workbook(wb)
  wb$clone()$setRowHeights(sheet, rows, heights)
}


#' Set worksheet column widths
#'
#' Set worksheet column widths to specific width or "auto".
#'
#' @param wb A [wbWorkbook] object
#' @param sheet A name or index of a worksheet
#' @param cols Indices of cols to set width
#' @param widths widths to set cols to specified in Excel column width units or "auto" for automatic sizing. The widths argument is
#' recycled to the length of cols. The default width is 8.43. Though there is no specific default width for Excel, it depends on
#' Excel version, operating system and DPI settings used. Setting it to specific value also is no guarantee that the output will be
#' of the selected width.
#' @param hidden Logical vector. If TRUE the column is hidden.
#' @details The global min and max column width for "auto" columns is set by (default values show):
#' \itemize{
#'   \item{options("openxlsx.minWidth" = 3)}
#'   \item{options("openxlsx.maxWidth" = 250)} ## This is the maximum width allowed in Excel
#' }
#'
#'   NOTE: The calculation of column widths can be slow for large worksheets.
#'
#'   NOTE: The `hidden` parameter may conflict with the one set in
#'   [wb_group_cols]; changing one will update the other.
#'
#' @export
#' @family workbook wrappers
#' @seealso [removeColWidths()]
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$addWorksheet("Sheet 1")
#'
#' ## set col widths
#' setColWidths(wb, 1, cols = c(1, 4, 6, 7, 9), widths = c(16, 15, 12, 18, 33))
#'
#' ## auto columns
#' wb$addWorksheet("Sheet 2")
#' writeData(wb, sheet = 2, x = iris)
#' setColWidths(wb, sheet = 2, cols = 1:5, widths = "auto")
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "setColWidthsExample.xlsx", overwrite = TRUE)
#' }
#'
setColWidths <- function(wb, sheet, cols, widths = 8.43, hidden = rep(FALSE, length(cols))) {
  assert_workbook(wb)
  sheet <- wb_validate_sheet(wb, sheet)

  # should do nothing if the cols' length is zero
  if (length(cols) == 0L) return(invisible(0))

  cols <- col2int(cols)

  if (length(widths) > length(cols)) {
    stop("More widths than columns supplied.")
  }

  if (length(hidden) > length(cols)) {
    stop("hidden argument is longer than cols.")
  }

  if (length(widths) < length(cols)) {
    widths <- rep(widths, length.out = length(cols))
  }

  if (length(hidden) < length(cols)) {
    hidden <- rep(hidden, length.out = length(cols))
  }

  # TODO add bestFit option?
  bestFit <- rep("1", length.out = length(cols))
  customWidth <- rep("1", length.out = length(cols))

  ## Remove duplicates
  widths <- widths[!duplicated(cols)]
  hidden <- hidden[!duplicated(cols)]
  cols <- cols[!duplicated(cols)]

  col_df <- wb$worksheets[[sheet]]$unfold_cols()

  if (any(widths == "auto")) {

    df <- wb_to_df(wb, sheet = sheet, cols = cols, colNames = FALSE)
    # TODO format(x) might not be the way it is formatted in the xlsx file.
    col_width <- vapply(df, function(x) {max(nchar(format(x)))}, NA_real_)

    # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column
    fw <- system.file("extdata", "fontwidth/FontWidth.csv", package = "openxlsx2")
    font_width_tab <- read.csv(fw)

    # TODO base font might not be the font used in this column
    base_font <- wb_get_base_font(wb)
    font <- base_font$name$val
    size <- as.integer(base_font$size$val)

    sel <- font_width_tab$FontFamilyName == font & font_width_tab$FontSize == size
    # maximum digit width of selected font
    mdw <- font_width_tab$Width[sel]

    # formula from openxml.spreadsheet.column documentation. The formula returns exactly the expected
    # value, but the output in excel is still off. Therefore round to create even numbers. In my tests
    # the results were close to the initial col_width sizes. Character width is still bad, numbers are
    # way larger, therefore characters cells are to wide. Not sure if we need improve this.
    widths <- trunc((col_width * mdw + 5) / mdw * 256) / 256
    widths <- round(widths)
  }

  # create empty cols
  if (NROW(col_df) == 0)
    col_df <- col_to_df(read_xml(wb$createCols(sheet, n = max(cols))))

  # found a few cols, but not all required cols. create the missing columns
  if (any(!cols %in% as.numeric(col_df$min))) {
    beg <- max(as.numeric(col_df$min)) + 1
    end <- max(cols)

    # new columns
    new_cols <- col_to_df(read_xml(wb$createCols(sheet, beg = beg, end = end)))

    # rbind only the missing columns. avoiding dups
    sel <- !new_cols$min %in% col_df$min
    col_df <- rbind(col_df, new_cols[sel,])
    col_df <- col_df[order(as.numeric(col_df[, "min"])),]
  }

  select <- as.numeric(col_df$min) %in% cols
  col_df$width[select] <- widths
  col_df$hidden[select] <- tolower(hidden)
  col_df$bestFit[select] <- bestFit
  col_df$customWidth[select] <- customWidth
  wb$worksheets[[sheet]]$fold_cols(col_df)

}

#' @name removeColWidths
#' @title Remove column widths from a worksheet

#' @description Remove column widths from a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Indices of columns to remove custom width (if any) from.
#' @seealso [setColWidths()]
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- loadWorkbook(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#'
#' ## remove column widths in columns 1 to 20
#' removeColWidths(wb, 1, cols = 1:20)
#' \dontrun{
#' wb_save(wb, "removeColWidthsExample.xlsx", overwrite = TRUE)
#' }
removeColWidths <- function(wb, sheet, cols) {
  sheet <- wb_validate_sheet(wb, sheet)
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  if (!is.numeric(cols)) {
    cols <- col2int(cols)
  }

  customCols <- as.integer(names(wb$colWidths[[sheet]]))
  removeInds <- which(customCols %in% cols)
  if (length(removeInds)) {
    remainingCols <- customCols[-removeInds]
    if (length(remainingCols) == 0) {
      wb$colWidths[[sheet]] <- list()
    } else {
      rem_widths <- wb$colWidths[[sheet]][-removeInds]
      names(rem_widths) <- as.character(remainingCols)
      wb$colWidths[[sheet]] <- rem_widths
    }
  }
}



#' @name removeRowHeights
#' @title Remove custom row heights from a worksheet
#' @description Remove row heights from a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Indices of rows to remove custom height (if any) from.
#' @seealso [wb_set_row_heights()]
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- loadWorkbook(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#'
#' ## remove any custom row heights in rows 1 to 10
#' removeRowHeights(wb, 1, rows = 1:10)
#' \dontrun{
#' wb_save(wb, "removeRowHeightsExample.xlsx", overwrite = TRUE)
#' }
removeRowHeights <- function(wb, sheet, rows) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  sheet <- wb_validate_sheet(wb, sheet)

  customRows <- as.integer(names(wb$rowHeights[[sheet]]))
  removeInds <- which(customRows %in% rows)
  if (length(removeInds)) {
    wb$rowHeights[[sheet]] <- wb$rowHeights[[sheet]][-removeInds]
  }
}


# images ------------------------------------------------------------------


#' @name insertPlot
#' @title Insert the current plot into a worksheet
#' @description The current plot is saved to a temporary image file using dev.copy.
#' This file is then written to the workbook using insertImage.
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param xy Alternate way to specify startRow and startCol.  A vector of length 2 of form (startcol, startRow)
#' @param startRow Row coordinate of upper left corner of figure. `xy[[2]]` when xy is given.
#' @param startCol Column coordinate of upper left corner of figure. `xy[[1]]` when xy is given.
#' @param width Width of figure. Defaults to 6in.
#' @param height Height of figure . Defaults to 4in.
#' @param fileType File type of image
#' @param units Units of width and height. Can be "in", "cm" or "px"
#' @param dpi Image resolution
#' @seealso [insertImage()]
#' @export
#' @examples
#' \dontrun{
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$addWorksheet("Sheet 1", showGridLines = FALSE)
#'
#' ## create plot objects
#' require(ggplot2)
#' p1 <- qplot(mpg,
#'   data = mtcars, geom = "density",
#'   fill = as.factor(gear), alpha = I(.5), main = "Distribution of Gas Mileage"
#' )
#' p2 <- qplot(age, circumference,
#'   data = Orange, geom = c("point", "line"), colour = Tree
#' )
#'
#' ## Insert currently displayed plot to sheet 1, row 1, column 1
#' print(p1) # plot needs to be showing
#' insertPlot(wb, 1, width = 5, height = 3.5, fileType = "png", units = "in")
#'
#' ## Insert plot 2
#' print(p2)
#' insertPlot(wb, 1, xy = c("J", 2), width = 16, height = 10, fileType = "png", units = "cm")
#'
#' ## Save workbook
#' wb_save(wb, "insertPlotExample.xlsx", overwrite = TRUE)
#' }
insertPlot <- function(wb, sheet, width = 6, height = 4, xy = NULL,
  startRow = 1, startCol = 1, fileType = "png", units = "in", dpi = 300) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  if (is.null(dev.list()[[1]])) {
    warning("No plot to insert.")
    return()
  }

  assert_workbook(wb)

  if (!is.null(xy)) {
    startCol <- xy[[1]]
    startRow <- xy[[2]]
  }

  fileType <- tolower(fileType)
  units <- tolower(units)

  if (fileType == "jpg") {
    fileType <- "jpeg"
  }

  if (!fileType %in% c("png", "jpeg", "tiff", "bmp")) {
    stop("Invalid file type.\nfileType must be one of: png, jpeg, tiff, bmp")
  }

  if (!units %in% c("cm", "in", "px")) {
    stop("Invalid units.\nunits must be one of: cm, in, px")
  }

  fileName <- tempfile(pattern = "figureImage", fileext = paste0(".", fileType))

  if (fileType == "bmp") {
    dev.copy(bmp, filename = fileName, width = width, height = height, units = units, res = dpi)
  } else if (fileType == "jpeg") {
    dev.copy(jpeg, filename = fileName, width = width, height = height, units = units, quality = 100, res = dpi)
  } else if (fileType == "png") {
    dev.copy(png, filename = fileName, width = width, height = height, units = units, res = dpi)
  } else if (fileType == "tiff") {
    dev.copy(tiff, filename = fileName, width = width, height = height, units = units, compression = "none", res = dpi)
  }

  ## write image
  invisible(dev.off())

  insertImage(wb = wb, sheet = sheet, file = fileName, width = width, height = height, startRow = startRow, startCol = startCol, units = units, dpi = dpi)
}


#' @title Remove a worksheet from a workbook
#' @description Remove a worksheet from a Workbook object
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @description Remove a worksheet from a workbook
#' @export
#' @examples
#' ## load a workbook
#' wb <- loadWorkbook(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#'
#' ## Remove sheet 2
#' wb <- wb_remove_worksheet(wb, 2)
#'
#' ## save the modified workbook
#' \dontrun{
#' wb_save(wb, "removeWorksheetExample.xlsx", overwrite = TRUE)
#' }
wb_remove_worksheet <- function(wb, sheet) {
  assert_workbook(wb)
  wb$clone()$removeWorksheet(sheet)
}


# base font ---------------------------------------------------------------

#' @name modifyBaseFont
#' @title Modify the default font
#' @description Modify the default font for this workbook
#' @param wb A workbook object
#' @param fontSize font size
#' @param fontColour font colour
#' @param fontName Name of a font
#' @details The font name is not validated in anyway.  Excel replaces unknown font names
#' with Arial. Base font is black, size 11, Calibri.
#' @export
#' @examples
#' ## create a workbook
#' wb <- wb_workbook()
#' wb$addWorksheet("S1")
#' ## modify base font to size 10 Arial Narrow in red
#' modifyBaseFont(wb, fontSize = 10, fontColour = "#FF0000", fontName = "Arial Narrow")
#'
#' writeData(wb, "S1", iris)
#' writeDataTable(wb, "S1", x = iris, startCol = 10) ## font colour does not affect tables
#' \dontrun{
#' wb_save(wb, "modifyBaseFontExample.xlsx", overwrite = TRUE)
#' }
modifyBaseFont <- function(wb, fontSize = 11, fontColour = "black", fontName = "Calibri") {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  assert_workbook(wb)

  if (fontSize < 0) stop("Invalid fontSize")
  fontColour <- validateColour(fontColour)

  wb$styles_mgr$styles$fonts[[1]] <- sprintf('<font><sz val="%s"/><color rgb="%s"/><name val="%s"/></font>', fontSize, fontColour, fontName)
}


#' Return the workbook default font
#'
#' Get the base font used in the workbook.
#' @param wb A [wbWorkbook] object
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' ## create a workbook
#' wb <- wb_workbook()
#' wb_get_base_font(wb)
#'
#' ## modify base font to size 10 Arial Narrow in red
#' modifyBaseFont(wb, fontSize = 10, fontColour = "#FF0000", fontName = "Arial Narrow")
#'
#' wb_get_base_font(wb)
wb_get_base_font <- function(wb) {
  # TODO all of these class checks need to be cleaned up
  assert_workbook(wb)
  wb$getBaseFont()
}

#' @name setHeaderFooter
#' @title Set document headers and footers
#' @description Set document headers and footers
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param header document header. Character vector of length 3 corresponding to positions left, center, right. Use NA to skip a position.
#' @param footer document footer. Character vector of length 3 corresponding to positions left, center, right. Use NA to skip a position.
#' @param evenHeader document header for even pages.
#' @param evenFooter document footer for even pages.
#' @param firstHeader document header for first page only.
#' @param firstFooter document footer for first page only.
#' @details Headers and footers can contain special tags
#' \itemize{
#'   \item{**&\[Page\]**}{ Page number}
#'   \item{**&\[Pages\]**}{ Number of pages}
#'   \item{**&\[Date\]**}{ Current date}
#'   \item{**&\[Time\]**}{ Current time}
#'   \item{**&\[Path\]**}{ File path}
#'   \item{**&\[File\]**}{ File name}
#'   \item{**&\[Tab\]**}{ Worksheet name}
#' }
#' @export
#' @seealso [wb_add_worksheet()] to set headers and footers when adding a worksheet
#' @examples
#' wb <- wb_workbook()
#'
#' wb$addWorksheet("S1")
#' wb$addWorksheet("S2")
#' wb$addWorksheet("S3")
#' wb$addWorksheet("S4")
#'
#' writeData(wb, 1, 1:400)
#' writeData(wb, 2, 1:400)
#' writeData(wb, 3, 3:400)
#' writeData(wb, 4, 3:400)
#'
#' setHeaderFooter(wb,
#'   sheet = "S1",
#'   header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
#'   footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
#'   evenHeader = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
#'   evenFooter = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
#'   firstHeader = c("TOP", "OF FIRST", "PAGE"),
#'   firstFooter = c("BOTTOM", "OF FIRST", "PAGE")
#' )
#'
#' setHeaderFooter(wb,
#'   sheet = 2,
#'   header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
#'   footer = c("&[Path]&[File]", NA, "&[Tab]"),
#'   firstHeader = c(NA, "Center Header of First Page", NA),
#'   firstFooter = c(NA, "Center Footer of First Page", NA)
#' )
#'
#' setHeaderFooter(wb,
#'   sheet = 3,
#'   header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
#'   footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2")
#' )
#'
#' setHeaderFooter(wb,
#'   sheet = 4,
#'   firstHeader = c("FIRST ONLY L", NA, "FIRST ONLY R"),
#'   firstFooter = c("FIRST ONLY L", NA, "FIRST ONLY R")
#' )
#' \dontrun{
#' wb_save(wb, "setHeaderFooterExample.xlsx", overwrite = TRUE)
#' }
setHeaderFooter <- function(wb, sheet,
  header = NULL,
  footer = NULL,
  evenHeader = NULL,
  evenFooter = NULL,
  firstHeader = NULL,
  firstFooter = NULL) {

  assert_workbook(wb)
  sheet <- wb_validate_sheet(wb, sheet)

  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  if (!is.null(header) & length(header) != 3) {
    stop("header must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(footer) & length(footer) != 3) {
    stop("footer must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(evenHeader) & length(evenHeader) != 3) {
    stop("evenHeader must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(evenFooter) & length(evenFooter) != 3) {
    stop("evenFooter must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(firstHeader) & length(firstHeader) != 3) {
    stop("firstHeader must have length 3 where elements correspond to positions: left, center, right.")
  }

  if (!is.null(firstFooter) & length(firstFooter) != 3) {
    stop("firstFooter must have length 3 where elements correspond to positions: left, center, right.")
  }

  oddHeader <- headerFooterSub(header)
  oddFooter <- headerFooterSub(footer)
  evenHeader <- headerFooterSub(evenHeader)
  evenFooter <- headerFooterSub(evenFooter)
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


  wb$worksheets[[sheet]]$headerFooter <- hf
}



#' @name pageSetup
#' @title Set page margins, orientation and print scaling
#' @description Set page margins, orientation and print scaling
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param orientation Page orientation. One of "portrait" or "landscape"
#' @param scale Print scaling. Numeric value between 10 and 400
#' @param left left page margin in inches
#' @param right right page margin in inches
#' @param top top page margin in inches
#' @param bottom bottom page margin in inches
#' @param header header margin in inches
#' @param footer footer margin in inches
#' @param fitToWidth If `TRUE`, worksheet is scaled to fit to page width on printing.
#' @param fitToHeight If `TRUE`, worksheet is scaled to fit to page height on printing.
#' @param paperSize See details. Default value is 9 (A4 paper).
#' @param printTitleRows Rows to repeat at top of page when printing. Integer vector.
#' @param printTitleCols Columns to repeat at left when printing. Integer vector.
#' @param summaryRow Location of summary rows in groupings. One of "Above" or "Below".
#' @param summaryCol Location of summary columns in groupings. One of "Right" or "Left".
#' @export
#' @details
#' paperSize is an integer corresponding to:
#' \itemize{
#' \item{**1**}{ Letter paper (8.5 in. by 11 in.)}
#' \item{**2**}{ Letter small paper (8.5 in. by 11 in.)}
#' \item{**3**}{ Tabloid paper (11 in. by 17 in.)}
#' \item{**4**}{ Ledger paper (17 in. by 11 in.)}
#' \item{**5**}{ Legal paper (8.5 in. by 14 in.)}
#' \item{**6**}{ Statement paper (5.5 in. by 8.5 in.)}
#' \item{**7**}{ Executive paper (7.25 in. by 10.5 in.)}
#' \item{**8**}{ A3 paper (297 mm by 420 mm)}
#' \item{**9**}{ A4 paper (210 mm by 297 mm)}
#' \item{**10**}{ A4 small paper (210 mm by 297 mm)}
#' \item{**11**}{ A5 paper (148 mm by 210 mm)}
#' \item{**12**}{ B4 paper (250 mm by 353 mm)}
#' \item{**13**}{ B5 paper (176 mm by 250 mm)}
#' \item{**14**}{ Folio paper (8.5 in. by 13 in.)}
#' \item{**15**}{ Quarto paper (215 mm by 275 mm)}
#' \item{**16**}{ Standard paper (10 in. by 14 in.)}
#' \item{**17**}{ Standard paper (11 in. by 17 in.)}
#' \item{**18**}{ Note paper (8.5 in. by 11 in.)}
#' \item{**19**}{ #9 envelope (3.875 in. by 8.875 in.)}
#' \item{**20**}{ #10 envelope (4.125 in. by 9.5 in.)}
#' \item{**21**}{ #11 envelope (4.5 in. by 10.375 in.)}
#' \item{**22**}{ #12 envelope (4.75 in. by 11 in.)}
#' \item{**23**}{ #14 envelope (5 in. by 11.5 in.)}
#' \item{**24**}{ C paper (17 in. by 22 in.)}
#' \item{**25**}{ D paper (22 in. by 34 in.)}
#' \item{**26**}{ E paper (34 in. by 44 in.)}
#' \item{**27**}{ DL envelope (110 mm by 220 mm)}
#' \item{**28**}{ C5 envelope (162 mm by 229 mm)}
#' \item{**29**}{ C3 envelope (324 mm by 458 mm)}
#' \item{**30**}{ C4 envelope (229 mm by 324 mm)}
#' \item{**31**}{ C6 envelope (114 mm by 162 mm)}
#' \item{**32**}{ C65 envelope (114 mm by 229 mm)}
#' \item{**33**}{ B4 envelope (250 mm by 353 mm)}
#' \item{**34**}{ B5 envelope (176 mm by 250 mm)}
#' \item{**35**}{ B6 envelope (176 mm by 125 mm)}
#' \item{**36**}{ Italy envelope (110 mm by 230 mm)}
#' \item{**37**}{ Monarch envelope (3.875 in. by 7.5 in.).}
#' \item{**38**}{ 6 3/4 envelope (3.625 in. by 6.5 in.)}
#' \item{**39**}{ US standard fanfold (14.875 in. by 11 in.)}
#' \item{**40**}{ German standard fanfold (8.5 in. by 12 in.)}
#' \item{**41**}{ German legal fanfold (8.5 in. by 13 in.)}
#' \item{**42**}{ ISO B4 (250 mm by 353 mm)}
#' \item{**43**}{ Japanese double postcard (200 mm by 148 mm)}
#' \item{**44**}{ Standard paper (9 in. by 11 in.)}
#' \item{**45**}{ Standard paper (10 in. by 11 in.)}
#' \item{**46**}{ Standard paper (15 in. by 11 in.)}
#' \item{**47**}{ Invite envelope (220 mm by 220 mm)}
#' \item{**50**}{ Letter extra paper (9.275 in. by 12 in.)}
#' \item{**51**}{ Legal extra paper (9.275 in. by 15 in.)}
#' \item{**52**}{ Tabloid extra paper (11.69 in. by 18 in.)}
#' \item{**53**}{ A4 extra paper (236 mm by 322 mm)}
#' \item{**54**}{ Letter transverse paper (8.275 in. by 11 in.)}
#' \item{**55**}{ A4 transverse paper (210 mm by 297 mm)}
#' \item{**56**}{ Letter extra transverse paper (9.275 in. by 12 in.)}
#' \item{**57**}{ SuperA/SuperA/A4 paper (227 mm by 356 mm)}
#' \item{**58**}{ SuperB/SuperB/A3 paper (305 mm by 487 mm)}
#' \item{**59**}{ Letter plus paper (8.5 in. by 12.69 in.)}
#' \item{**60**}{ A4 plus paper (210 mm by 330 mm)}
#' \item{**61**}{ A5 transverse paper (148 mm by 210 mm)}
#' \item{**62**}{ JIS B5 transverse paper (182 mm by 257 mm)}
#' \item{**63**}{ A3 extra paper (322 mm by 445 mm)}
#' \item{**64**}{ A5 extra paper (174 mm by 235 mm)}
#' \item{**65**}{ ISO B5 extra paper (201 mm by 276 mm)}
#' \item{**66**}{ A2 paper (420 mm by 594 mm)}
#' \item{**67**}{ A3 transverse paper (297 mm by 420 mm)}
#' \item{**68**}{ A3 extra transverse paper (322 mm by 445 mm)}
#' }
#' @examples
#' wb <- wb_workbook()
#' wb$addWorksheet("S1")
#' wb$addWorksheet("S2")
#' writeDataTable(wb, 1, x = iris[1:30, ])
#' writeDataTable(wb, 2, x = iris[1:30, ], xy = c("C", 5))
#'
#' ## landscape page scaled to 50%
#' pageSetup(wb, sheet = 1, orientation = "landscape", scale = 50)
#'
#' ## portrait page scales to 300% with 0.5in left and right margins
#' pageSetup(wb, sheet = 2, orientation = "portrait", scale = 300, left = 0.5, right = 0.5)
#'
#'
#' ## print titles
#' wb$addWorksheet("print_title_rows")
#' wb$addWorksheet("print_title_cols")
#'
#' writeData(wb, "print_title_rows", rbind(iris, iris, iris, iris))
#' writeData(wb, "print_title_cols", x = rbind(mtcars, mtcars, mtcars), rowNames = TRUE)
#'
#' pageSetup(wb, sheet = "print_title_rows", printTitleRows = 1) ## first row
#' pageSetup(wb, sheet = "print_title_cols", printTitleCols = 1, printTitleRows = 1)
#' \dontrun{
#' wb_save(wb, "pageSetupExample.xlsx", overwrite = TRUE)
#' }
pageSetup <- function(wb, sheet, orientation = NULL, scale = 100,
  left = 0.7, right = 0.7, top = 0.75, bottom = 0.75,
  header = 0.3, footer = 0.3,
  fitToWidth = FALSE, fitToHeight = FALSE, paperSize = NULL,
  printTitleRows = NULL, printTitleCols = NULL,
  summaryRow = NULL, summaryCol = NULL) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  assert_workbook(wb)

  sheet <- wb_validate_sheet(wb, sheet)
  xml <- wb$worksheets[[sheet]]$pageSetup

  if (!is.null(orientation)) {
    orientation <- tolower(orientation)
    if (!orientation %in% c("portrait", "landscape")) stop("Invalid page orientation.")
  } else {
    # if length(xml) == 1 then use if () {} else {}
    orientation <- ifelse(grepl("landscape", xml), "landscape", "portrait") ## get existing
  }

  if (scale < 10 | scale > 400) {
    stop("Scale must be between 10 and 400.")
  }

  if (!is.null(paperSize)) {
    paperSizes <- 1:68
    paperSizes <- paperSizes[!paperSizes %in% 48:49]
    if (!paperSize %in% paperSizes) {
      stop("paperSize must be an integer in range [1, 68]. See ?pageSetup details.")
    }
    paperSize <- as.integer(paperSize)
  } else {
    paperSize <- regmatches(xml, regexpr('(?<=paperSize=")[0-9]+', xml, perl = TRUE)) ## get existing
  }


  ##############################
  ## Keep defaults on orientation, hdpi, vdpi, paperSize
  hdpi <- regmatches(xml, regexpr('(?<=horizontalDpi=")[0-9]+', xml, perl = TRUE))
  vdpi <- regmatches(xml, regexpr('(?<=verticalDpi=")[0-9]+', xml, perl = TRUE))


  ##############################
  ## Update
  wb$worksheets[[sheet]]$pageSetup <- sprintf(
    '<pageSetup paperSize="%s" orientation="%s" scale = "%s" fitToWidth="%s" fitToHeight="%s" horizontalDpi="%s" verticalDpi="%s" r:id="rId2"/>',
    paperSize, orientation, scale, as.integer(fitToWidth), as.integer(fitToHeight), hdpi, vdpi
  )

  if (fitToHeight | fitToWidth) {
    wb$worksheets[[sheet]]$sheetPr <- unique(c(wb$worksheets[[sheet]]$sheetPr, '<pageSetUpPr fitToPage="1"/>'))
  }

  wb$worksheets[[sheet]]$pageMargins <-
    sprintf('<pageMargins left="%s" right="%s" top="%s" bottom="%s" header="%s" footer="%s"/>', left, right, top, bottom, header, footer)

  validRow <- function(summaryRow) {
    return(tolower(summaryRow) %in% c("above", "below"))
  }
  validCol <- function(summaryCol) {
    return(tolower(summaryCol) %in% c("left", "right"))
  }

  outlinepr <- ""

  if (!is.null(summaryRow)) {

    if (!validRow(summaryRow)) {
      stop("Invalid \`summaryRow\` option. Must be one of \"Above\" or \"Below\".")
    } else if (tolower(summaryRow) == "above") {
      outlinepr <- ' summaryBelow=\"0\"'
    } else {
      outlinepr <- ' summaryBelow=\"1\"'
    }
  }

  if (!is.null(summaryCol)) {

    if (!validCol(summaryCol)) {
      stop("Invalid \`summaryCol\` option. Must be one of \"Left\" or \"Right\".")
    } else if (tolower(summaryCol) == "left") {
      outlinepr <- paste0(outlinepr, ' summaryRight=\"0\"')
    } else {
      outlinepr <- paste0(outlinepr, ' summaryRight=\"1\"')
    }
  }

  if (!stri_isempty(outlinepr)) {
    wb$worksheets[[sheet]]$sheetPr <- unique(c(wb$worksheets[[sheet]]$sheetPr, paste0("<outlinePr", outlinepr, "/>")))
  }

  ## print Titles
  if (!is.null(printTitleRows) & is.null(printTitleCols)) {
    if (!is.numeric(printTitleRows)) {
      stop("printTitleRows must be numeric.")
    }

    wb$createNamedRegion(
      ref1 = paste0("$", min(printTitleRows)),
      ref2 = paste0("$", max(printTitleRows)),
      name = "_xlnm.Print_Titles",
      sheet = names(wb)[[sheet]],
      localSheetId = sheet - 1L
    )
  } else if (!is.null(printTitleCols) & is.null(printTitleRows)) {
    if (!is.numeric(printTitleCols)) {
      stop("printTitleCols must be numeric.")
    }

    cols <- int2col(range(printTitleCols))
    wb$createNamedRegion(
      ref1 = paste0("$", cols[1]),
      ref2 = paste0("$", cols[2]),
      name = "_xlnm.Print_Titles",
      sheet = names(wb)[[sheet]],
      localSheetId = sheet - 1L
    )
  } else if (!is.null(printTitleCols) & !is.null(printTitleRows)) {
    if (!is.numeric(printTitleRows)) {
      stop("printTitleRows must be numeric.")
    }

    if (!is.numeric(printTitleCols)) {
      stop("printTitleCols must be numeric.")
    }

    cols <- int2col(range(printTitleCols))
    rows <- range(printTitleRows)

    cols <- paste(paste0("$", cols[1]), paste0("$", cols[2]), sep = ":")
    rows <- paste(paste0("$", rows[1]), paste0("$", rows[2]), sep = ":")
    localSheetId <- sheet - 1L
    sheet <- names(wb)[[sheet]]

    wb$workbook$definedNames <- c(
      wb$workbook$definedNames,
      sprintf('<definedName name="_xlnm.Print_Titles" localSheetId="%s">\'%s\'!%s,\'%s\'!%s</definedName>', localSheetId, sheet, cols, sheet, rows)
    )
  }
}


# protect -----------------------------------------------------------------

#' @name protectWorksheet
#' @title Protect a worksheet from modifications
#' @description Protect or unprotect a worksheet from modifications by the user in the graphical user interface. Replaces an existing protection.
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param protect Whether to protect or unprotect the sheet (default=TRUE)
#' @param password (optional) password required to unprotect the worksheet
#' @param lockSelectingLockedCells Whether selecting locked cells is locked
#' @param lockSelectingUnlockedCells Whether selecting unlocked cells is locked
#' @param lockFormattingCells Whether formatting cells is locked
#' @param lockFormattingColumns Whether formatting columns is locked
#' @param lockFormattingRows Whether formatting rows is locked
#' @param lockInsertingColumns Whether inserting columns is locked
#' @param lockInsertingRows Whether inserting rows is locked
#' @param lockInsertingHyperlinks Whether inserting hyperlinks is locked
#' @param lockDeletingColumns Whether deleting columns is locked
#' @param lockDeletingRows Whether deleting rows is locked
#' @param lockSorting Whether sorting is locked
#' @param lockAutoFilter Whether auto-filter is locked
#' @param lockPivotTables Whether pivot tables are locked
#' @param lockObjects Whether objects are locked
#' @param lockScenarios Whether scenarios are locked
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$addWorksheet("S1")
#' writeDataTable(wb, 1, x = iris[1:30, ])
#' # Formatting cells / columns is allowed , but inserting / deleting columns is protected:
#' protectWorksheet(wb, "S1",
#'   protect = TRUE,
#'   lockFormattingCells = FALSE, lockFormattingColumns = FALSE,
#'   lockInsertingColumns = TRUE, lockDeletingColumns = TRUE
#' )
#'
#' # Remove the protection
#' protectWorksheet(wb, "S1", protect = FALSE)
#' \dontrun{
#' wb_save(wb, "pageSetupExample.xlsx", overwrite = TRUE)
#' }
protectWorksheet <- function(wb, sheet, protect = TRUE, password = NULL,
  lockSelectingLockedCells = NULL, lockSelectingUnlockedCells = NULL,
  lockFormattingCells = NULL, lockFormattingColumns = NULL, lockFormattingRows = NULL,
  lockInsertingColumns = NULL, lockInsertingRows = NULL, lockInsertingHyperlinks = NULL,
  lockDeletingColumns = NULL, lockDeletingRows = NULL,
  lockSorting = NULL, lockAutoFilter = NULL, lockPivotTables = NULL,
  lockObjects = NULL, lockScenarios = NULL) {

  assert_workbook(wb)

  sheet <- wb_validate_sheet(wb, sheet)

  props <- c()

  if (!missing(password) && !is.null(password)) {
    props["password"] <- hashPassword(password)
  }

  if (!missing(lockSelectingLockedCells) && !is.null(lockSelectingLockedCells)) {
    props["selectLockedCells"] <- toString(as.numeric(lockSelectingLockedCells))
  }
  if (!missing(lockSelectingUnlockedCells) && !is.null(lockSelectingUnlockedCells)) {
    props["selectUnlockedCells"] <- toString(as.numeric(lockSelectingUnlockedCells))
  }
  if (!missing(lockFormattingCells) && !is.null(lockFormattingCells)) {
    props["formatCells"] <- toString(as.numeric(lockFormattingCells))
  }
  if (!missing(lockFormattingColumns) && !is.null(lockFormattingColumns)) {
    props["formatColumns"] <- toString(as.numeric(lockFormattingColumns))
  }
  if (!missing(lockFormattingRows) && !is.null(lockFormattingRows)) {
    props["formatRows"] <- toString(as.numeric(lockFormattingRows))
  }
  if (!missing(lockInsertingColumns) && !is.null(lockInsertingColumns)) {
    props["insertColumns"] <- toString(as.numeric(lockInsertingColumns))
  }
  if (!missing(lockInsertingRows) && !is.null(lockInsertingRows)) {
    props["insertRows"] <- toString(as.numeric(lockInsertingRows))
  }
  if (!missing(lockInsertingHyperlinks) && !is.null(lockInsertingHyperlinks)) {
    props["insertHyperlinks"] <- toString(as.numeric(lockInsertingHyperlinks))
  }
  if (!missing(lockDeletingColumns) && !is.null(lockDeletingColumns)) {
    props["deleteColumns"] <- toString(as.numeric(lockDeletingColumns))
  }
  if (!missing(lockDeletingRows) && !is.null(lockDeletingRows)) {
    props["deleteRows"] <- toString(as.numeric(lockDeletingRows))
  }
  if (!missing(lockSorting) && !is.null(lockSorting)) {
    props["sort"] <- toString(as.numeric(lockSorting))
  }
  if (!missing(lockAutoFilter) && !is.null(lockAutoFilter)) {
    props["autoFilter"] <- toString(as.numeric(lockAutoFilter))
  }
  if (!missing(lockPivotTables) && !is.null(lockPivotTables)) {
    props["pivotTables"] <- toString(as.numeric(lockPivotTables))
  }
  if (!missing(lockObjects) && !is.null(lockObjects)) {
    props["objects"] <- toString(as.numeric(lockObjects))
  }
  if (!missing(lockScenarios) && !is.null(lockScenarios)) {
    props["scenarios"] <- toString(as.numeric(lockScenarios))
  }

  if (protect) {
    props["sheet"] <- "1"
    wb$worksheets[[sheet]]$sheetProtection <- sprintf("<sheetProtection %s/>", paste(names(props), '="', props, '"', collapse = " ", sep = ""))
  } else {
    wb$worksheets[[sheet]]$sheetProtection <- ""
  }
}


#' @name protectWorkbook
#' @title Protect a workbook from modifications
#' @description Protect or unprotect a workbook from modifications by the user in the graphical user interface. Replaces an existing protection.
#' @param wb A workbook object
#' @param protect Whether to protect or unprotect the sheet (default=TRUE)
#' @param password (optional) password required to unprotect the workbook
#' @param lockStructure Whether the workbook structure should be locked
#' @param lockWindows Whether the window position of the spreadsheet should be locked
#' @param type Lock type, default 1 - xlsx with password. 2 - xlsx recommends read-only. 4 - xlsx enforces read-only. 8 - xlsx is locked for annotation.
#' @param fileSharing Whether to enable a popup requesting the unlock password is prompted
#' @param username The username for the fileSharing popup
#' @param readOnlyRecommended Whether or not a post unlock message appears stating that the workbook is recommended to be opened in readonly mode.
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$addWorksheet("S1")
#' protectWorkbook(wb, protect = TRUE, password = "Password", lockStructure = TRUE)
#' \dontrun{
#' wb_save(wb, "WorkBook_Protection.xlsx", overwrite = TRUE)
#' }
#' # Remove the protection
#' protectWorkbook(wb, protect = FALSE)
#' \dontrun{
#' wb_save(wb, "WorkBook_Protection_unprotected.xlsx", overwrite = TRUE)
#' }
#' 
#' protectWorkbook(wb, protect = TRUE, password = "Password", lockStructure = TRUE, type = 2L, fileSharing = TRUE, username = "Test", readOnlyRecommended = TRUE)
#' 
protectWorkbook <- function(wb, protect = TRUE, password = NULL, lockStructure = FALSE, lockWindows = FALSE, type = 1L, fileSharing = FALSE, username = unname(Sys.info()["user"]), readOnlyRecommended = FALSE) {
  assert_workbook(wb)
  invisible(wb$protectWorkbook(protect = protect, password = password, lockStructure = lockStructure, lockWindows = lockWindows, type = type, fileSharing = fileSharing, username = username, readOnlyRecommended = readOnlyRecommended))
}


# grid lines --------------------------------------------------------------

#' @name showGridLines
#' @title Set worksheet gridlines to show or hide.
#' @description Set worksheet gridlines to show or hide.
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param showGridLines A logical. If `FALSE`, grid lines are hidden.
#' @export
#' @examples
#' wb <- loadWorkbook(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#' names(wb) ## list worksheets in workbook
#' showGridLines(wb, 1, showGridLines = FALSE)
#' showGridLines(wb, "testing", showGridLines = FALSE)
#' \dontrun{
#' wb_save(wb, "showGridLinesExample.xlsx", overwrite = TRUE)
#' }
showGridLines <- function(wb, sheet, showGridLines = FALSE) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  assert_workbook(wb)

  sheet <- wb_validate_sheet(wb, sheet)

  if (!is.logical(showGridLines)) stop("showGridLines must be a logical")


  sv <- wb$worksheets[[sheet]]$sheetViews
  showGridLines <- as.integer(showGridLines)
  ## If attribute exists gsub
  if (grepl("showGridLines", sv)) {
    sv <- gsub('showGridLines=".?[^"]', sprintf('showGridLines="%s', showGridLines), sv, perl = TRUE)
  } else {
    sv <- gsub("<sheetView ", sprintf('<sheetView showGridLines="%s" ', showGridLines), sv)
  }

  wb$worksheets[[sheet]]$sheetViews <- sv
}

# TODO hide gridlines?

# worksheet order ---------------------------------------------------------

#' @name worksheetOrder
#' @title Order of worksheets in xlsx file
#' @description Get/set order of worksheets in a Workbook object
#' @details This function does not reorder the worksheets within the workbook object, it simply
#' shuffles the order when writing to file.
#' @export
#' @examples
#' ## setup a workbook with 3 worksheets
#' wb <- wb_workbook()
#' wb$addWorksheet(sheet = "Sheet 1", showGridLines = FALSE)
#' writeDataTable(wb = wb, sheet = 1, x = iris)
#'
#' wb$addWorksheet(sheet = "mtcars (Sheet 2)", showGridLines = FALSE)
#' writeData(wb = wb, sheet = 2, x = mtcars)
#'
#' wb$addWorksheet(sheet = "Sheet 3", showGridLines = FALSE)
#' writeData(wb = wb, sheet = 3, x = Formaldehyde)
#'
#' worksheetOrder(wb)
#' names(wb)
#' worksheetOrder(wb) <- c(1, 3, 2) # switch position of sheets 2 & 3
#' writeData(wb, 2, 'This is still the "mtcars" worksheet', startCol = 15)
#' worksheetOrder(wb)
#' names(wb) ## ordering within workbook is not changed
#' \dontrun{
#' wb_save(wb, "worksheetOrderExample.xlsx", overwrite = TRUE)
#' }
#' worksheetOrder(wb) <- c(3, 2, 1)
#' \dontrun{
#' wb_save(wb, "worksheetOrderExample2.xlsx", overwrite = TRUE)
#' }
worksheetOrder <- function(wb) {
  assert_workbook(wb)
  wb$sheetOrder
}

#' @rdname worksheetOrder
#' @param wb A workbook object
#' @param value Vector specifying order to write worksheets to file
#' @export
`worksheetOrder<-` <- function(wb, value) {
  assert_workbook(wb)

  if (any(value != as.integer(value))) {
    stop("values must be integers")
  }

  value <- as.integer(value)

  value <- unique(value)
  if (length(value) != length(wb$worksheets)) {
    stop(sprintf("Worksheet order must be same length as number of worksheets [%s]", length(wb$worksheets)))
  }

  if (any(value > length(wb$worksheets))) {
    stop("Elements of order are greater than the number of worksheets")
  }

  wb$sheetOrder <- value

  invisible(wb)
}

#' @name createNamedRegion
#' @title Create / delete a named region
#' @description Create / delete a named region
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Numeric vector specifying rows to include in region
#' @param cols Numeric vector specifying columns to include in region
#' @param name Name for region. A character vector of length 1. Note region names musts be case-insensitive unique.
#' @param overwrite Boolean. Overwrite if exists? Default to FALSE
#' @details Region is given by: min(cols):max(cols) X min(rows):max(rows)
#' @seealso [getNamedRegions()] [deleteNamedRegion()]
#' @examples
#' ## create named regions
#' wb <- wb_workbook()
#' wb$addWorksheet("Sheet 1")
#'
#' ## specify region
#' writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
#' createNamedRegion(
#'   wb = wb,
#'   sheet = 1,
#'   name = "iris",
#'   rows = seq_len(nrow(iris) + 1),
#'   cols = seq_along(iris)
#' )
#'
#'
#' ## using writeData 'name' argument
#' writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
#'
#' out_file <- tempfile(fileext = ".xlsx")
#' \dontrun{
#' wb_save(wb, out_file, overwrite = TRUE)
#'
#' ## see named regions
#' getNamedRegions(wb) ## From Workbook object
#' getNamedRegions(out_file) ## From xlsx file
#'
#' ## delete one
#' deleteNamedRegion(wb = wb, name = "iris2")
#' getNamedRegions(wb)
#'
#' ## read named regions
#' df <- read.xlsx(wb, namedRegion = "iris")
#' head(df)
#'
#' df <- read.xlsx(out_file, namedRegion = "iris2")
#' head(df)
#' }
#' @rdname NamedRegion
#' @export
createNamedRegion <- function(wb, sheet, cols, rows, name, overwrite = FALSE) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  sheet <- wb_validate_sheet(wb, sheet)

  assert_workbook(wb)

  if (!is.numeric(rows)) {
    stop("rows argument must be a numeric/integer vector")
  }

  if (!is.numeric(cols)) {
    stop("cols argument must be a numeric/integer vector")
  }

  ## check name doesn't already exist
  ## named region

  ex_names <- regmatches(wb$workbook$definedNames, regexpr('(?<=name=")[^"]+', wb$workbook$definedNames, perl = TRUE))
  ex_names <- tolower(replaceXMLEntities(ex_names))

  if (tolower(name) %in% ex_names) {
    if (overwrite)
      wb$workbook$definedNames <- wb$workbook$definedNames[!ex_names %in% tolower(name)]
    else
      stop(sprintf("Named region with name '%s' already exists! Use overwrite  = TRUE if you want to replace it", name))
  } else if (grepl("^[A-Z]{1,3}[0-9]+$", name)) {
    stop("name cannot look like a cell reference.")
  }


  cols <- round(cols)
  rows <- round(rows)

  startCol <- min(cols)
  endCol <- max(cols)

  startRow <- min(rows)
  endRow <- max(rows)

  ref1 <- paste0("$", int2col(startCol), "$", startRow)
  ref2 <- paste0("$", int2col(endCol), "$", endRow)

  invisible(
    wb$createNamedRegion(ref1 = ref1, ref2 = ref2, name = name, sheet = wb$sheet_names[sheet])
  )
}

#' @export
#' @rdname NamedRegion
deleteNamedRegion <- function(wb, sheet, name) {

  assert_workbook(wb)

  # get all nown defined names
  dn <- getNamedRegions(wb)

  if (missing(name) & !missing(sheet)) {
    sheet <- wb_validate_sheet(wb, sheet)
    del <- dn$id[dn$sheet == sheet]
  } else if (!missing(name) & missing(sheet)) {
    del <- dn$id[dn$name == name]
  } else {
    sheet <- wb_validate_sheet(wb, sheet)
    del <- dn$id[dn$sheet == sheet & dn$name == name]
  }

  if (length(del)) {
    wb$workbook$definedNames <- wb$workbook$definedNames[-del]
  } else {
    if (!missing(name))
      warning(sprintf("Cannot find named region with name '%s'", name))
    # do not warn if wb and sheet are selected. deleteNamedRegion is 
    # called with every wb_remove_worksheet and would throw meaningless
    # warnings. For now simply assume if no name is defined, that the
    # user does not care, as long as no defined name remains on a sheet.
  }

  invisible(0)
}

# filters -----------------------------------------------------------------

#' @name addFilter
#' @title Add column filters
#' @description Add excel column filters to a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols columns to add filter to.
#' @param rows A row number.
#' @seealso [writeData()]
#' @details adds filters to worksheet columns, same as filter parameters in writeData.
#' writeDataTable automatically adds filters to first row of a table.
#' NOTE Can only have a single filter per worksheet unless using tables.
#' @export
#' @seealso [addFilter()]
#' @examples
#' wb <- wb_workbook()
#' wb$addWorksheet("Sheet 1")
#' wb$addWorksheet("Sheet 2")
#' wb$addWorksheet("Sheet 3")
#'
#' writeData(wb, 1, iris)
#' addFilter(wb, 1, row = 1, cols = seq_along(iris))
#'
#' ## Equivalently
#' writeData(wb, 2, x = iris, withFilter = TRUE)
#'
#' ## Similarly
#' writeDataTable(wb, 3, iris)
#' \dontrun{
#' wb_save(wb, path = "addFilterExample.xlsx", overwrite = TRUE)
#' }
addFilter <- function(wb, sheet, rows, cols) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  assert_workbook(wb)

  sheet <- wb_validate_sheet(wb, sheet)

  if (length(rows) != 1) {
    stop("row must be a numeric of length 1.")
  }

  if (!is.numeric(cols)) {
    cols <- col2int(cols)
  }

  wb$worksheets[[sheet]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', paste(getCellRefs(data.frame("x" = c(rows, rows), "y" = c(min(cols), max(cols)))), collapse = ":"))

  invisible(wb)
}

#' @name removeFilter
#' @title Remove a worksheet filter
#' @description Removes filters from addFilter() and writeData()
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$addWorksheet("Sheet 1")
#' wb$addWorksheet("Sheet 2")
#' wb$addWorksheet("Sheet 3")
#'
#' writeData(wb, 1, iris)
#' addFilter(wb, 1, row = 1, cols = seq_along(iris))
#'
#' ## Equivalently
#' writeData(wb, 2, x = iris, withFilter = TRUE)
#'
#' ## Similarly
#' writeDataTable(wb, 3, iris)
#'
#' ## remove filters
#' removeFilter(wb, 1:2) ## remove filters
#' removeFilter(wb, 3) ## Does not affect tables!
#' \dontrun{
#' wb_save(wb, path = "removeFilterExample.xlsx", overwrite = TRUE)
#' }
removeFilter <- function(wb, sheet) {
  assert_workbook(wb)

  for (s in sheet) {
    s <- wb_validate_sheet(wb, s)
    wb$worksheets[[s]]$autoFilter <- character()
  }

  invisible(wb)
}


# validations -------------------------------------------------------------

#' @name dataValidation
#' @title Add data validation to cells
#' @description Add Excel data validation to cells
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Contiguous columns to apply conditional formatting to
#' @param rows Contiguous rows to apply conditional formatting to
#' @param type One of 'whole', 'decimal', 'date', 'time', 'textLength', 'list' (see examples)
#' @param operator One of 'between', 'notBetween', 'equal',
#'  'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'
#' @param value a vector of length 1 or 2 depending on operator (see examples)
#' @param allowBlank logical
#' @param showInputMsg logical
#' @param showErrorMsg logical
#' @export
#' @examples
#' \dontrun{
#' wb <- wb_workbook()
#' wb$addWorksheet("Sheet 1")
#' wb$addWorksheet("Sheet 2")
#'
#' writeDataTable(wb, 1, x = iris[1:30, ])
#'
#' dataValidation(wb, 1,
#'   col = 1:3, rows = 2:31, type = "whole",
#'   operator = "between", value = c(1, 9)
#' )
#'
#' dataValidation(wb, 1,
#'   col = 5, rows = 2:31, type = "textLength",
#'   operator = "between", value = c(4, 6)
#' )
#'
#'
#' ## Date and Time cell validation
#' df <- data.frame(
#'   "d" = as.Date("2016-01-01") + -5:5,
#'   "t" = as.POSIXct("2016-01-01") + -5:5 * 10000
#' )
#'
#' writeData(wb, 2, x = df)
#' dataValidation(wb, 2,
#'   col = 1, rows = 2:12, type = "date",
#'   operator = "greaterThanOrEqual", value = as.Date("2016-01-01")
#' )
#'
#' dataValidation(wb, 2,
#'   col = 2, rows = 2:12, type = "time",
#'   operator = "between", value = df$t[c(4, 8)]
#' )
#' wb_save(wb, "dataValidationExample.xlsx", overwrite = TRUE)
#'
#'
#' ######################################################################
#' ## If type == 'list'
#' # operator argument is ignored.
#'
#' wb <- wb_workbook()
#' wb$addWorksheet("Sheet 1")
#' wb$addWorksheet("Sheet 2")
#'
#' writeDataTable(wb, sheet = 1, x = iris[1:30, ])
#' writeData(wb, sheet = 2, x = sample(iris$Sepal.Length, 10))
#'
#' dataValidation(wb, 1, col = 1, rows = 2:31, type = "list", value = "'Sheet 2'!$A$1:$A$10")
#' }
#'
#' # openXL(wb)
dataValidation <- function(wb, sheet, cols, rows, type, operator, value, allowBlank = TRUE, showInputMsg = TRUE, showErrorMsg = TRUE) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  ## rows and cols
  if (!is.numeric(cols)) {
    cols <- col2int(cols)
  }
  rows <- as.integer(rows)

  ## check length of value
  if (length(value) > 2) {
    stop("value argument must be length < 2")
  }

  valid_types <- c(
    "whole",
    "decimal",
    "date",
    "time", ## need to conv
    "textLength",
    "list"
  )

  if (!tolower(type) %in% tolower(valid_types)) {
    stop("Invalid 'type' argument!")
  }


  ## operator == 'between' we leave out
  valid_operators <- c(
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual"
  )

  if (tolower(type) != "list") {
    if (!tolower(operator) %in% tolower(valid_operators)) {
      stop("Invalid 'operator' argument!")
    }

    operator <- valid_operators[tolower(valid_operators) %in% tolower(operator)][1]
  } else {
    operator <- "between" ## ignored
  }

  if (!is.logical(allowBlank)) {
    stop("Argument 'allowBlank' musts be logical!")
  }

  if (!is.logical(showInputMsg)) {
    stop("Argument 'showInputMsg' musts be logical!")
  }

  if (!is.logical(showErrorMsg)) {
    stop("Argument 'showErrorMsg' musts be logical!")
  }

  ## All inputs validated

  type <- valid_types[tolower(valid_types) %in% tolower(type)][1]

  ## check input combinations
  if (type == "date" & !inherits(value, "Date")) {
    stop("If type == 'date' value argument must be a Date vector.")
  }

  if (type == "time" & !inherits(value, c("POSIXct", "POSIXt"))) {
    stop("If type == 'date' value argument must be a POSIXct or POSIXlt vector.")
  }


  value <- head(value, 2)
  allowBlank <- as.integer(allowBlank[1])
  showInputMsg <- as.integer(showInputMsg[1])
  showErrorMsg <- as.integer(showErrorMsg[1])

  if (type == "list") {
    invisible(wb$dataValidation_list(
      sheet = sheet,
      startRow = min(rows),
      endRow = max(rows),
      startCol = min(cols),
      endCol = max(cols),
      value = value,
      allowBlank = allowBlank,
      showInputMsg = showInputMsg,
      showErrorMsg = showErrorMsg
    ))
  } else {
    invisible(wb$dataValidation(
      sheet = sheet,
      startRow = min(rows),
      endRow = max(rows),
      startCol = min(cols),
      endCol = max(cols),
      type = type,
      operator = operator,
      value = value,
      allowBlank = allowBlank,
      showInputMsg = showInputMsg,
      showErrorMsg = showErrorMsg
    ))
  }



  invisible(0)
}


# visibility --------------------------------------------------------------

#' @name sheetVisibility
#' @title Get/set worksheet visible state
#' @description Get and set worksheet visible state
#' @param wb A workbook object
#' @return Character vector of worksheet names.
#' @return  Vector of "hidden", "visible", "veryHidden"
#' @examples
#'
#' wb <- wb_workbook()
#' wb$addWorksheet(sheet = "S1", visible = FALSE)
#' wb$addWorksheet(sheet = "S2", visible = TRUE)
#' wb$addWorksheet(sheet = "S3", visible = FALSE)
#'
#' sheetVisibility(wb)
#' sheetVisibility(wb)[1] <- TRUE ## show sheet 1
#' sheetVisibility(wb)[2] <- FALSE ## hide sheet 2
#' sheetVisibility(wb)[3] <- "hidden" ## hide sheet 3
#' sheetVisibility(wb)[3] <- "veryHidden" ## hide sheet 3 from UI
#' @export
sheetVisibility <- function(wb) {
  assert_workbook(wb)

  state <- rep("visible", length(wb$workbook$sheets))
  state[grepl("hidden", wb$workbook$sheets)] <- "hidden"
  state[grepl("veryHidden", wb$workbook$sheets, ignore.case = TRUE)] <- "veryHidden"


  return(state)
}

#' @rdname sheetVisibility
#' @param value a logical/character vector the same length as sheetVisibility(wb)
#' @export
`sheetVisibility<-` <- function(wb, value) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  value <- tolower(as.character(value))
  if (!any(value %in% c("true", "visible"))) {
    stop("A workbook must have atleast 1 visible worksheet.")
  }

  value[value %in% "true"] <- "visible"
  value[value %in% "false"] <- "hidden"
  value[value %in% "veryhidden"] <- "veryHidden"


  exState0 <- regmatches(wb$workbook$sheets, regexpr('(?<=state=")[^"]+', wb$workbook$sheets, perl = TRUE))
  exState <- tolower(exState0)
  exState[exState %in% "true"] <- "visible"
  exState[exState %in% "hidden"] <- "hidden"
  exState[exState %in% "false"] <- "hidden"
  exState[exState %in% "veryhidden"] <- "veryHidden"

  if (length(value) != length(wb$workbook$sheets)) {
    stop(sprintf("value vector must have length equal to number of worksheets in Workbook [%s]", length(exState)))
  }

  inds <- which(value != exState)
  if (length(inds) == 0) {
    return(invisible(wb))
  }

  for (i in seq_along(wb$worksheets)) {
    wb$workbook$sheets[i] <- gsub(exState0[i], value[i], wb$workbook$sheets[i], fixed = TRUE)
  }

  invisible(wb)
}


#' @name pageBreak
#' @title add a page break to a worksheet
#' @description insert page breaks into a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param i row or column number to insert page break.
#' @param type One of "row" or "column" for a row break or column break.
#' @export
#' @seealso [wb_add_worksheet()]
#' @examples
#' wb <- wb_workbook()
#' wb$addWorksheet("Sheet 1")
#' writeData(wb, sheet = 1, x = iris)
#'
#' pageBreak(wb, sheet = 1, i = 10, type = "row")
#' pageBreak(wb, sheet = 1, i = 20, type = "row")
#' pageBreak(wb, sheet = 1, i = 2, type = "column")
#' \dontrun{
#' wb_save(wb, "pageBreakExample.xlsx", TRUE)
#' }
#' ## In Excel: View tab -> Page Break Preview
pageBreak <- function(wb, sheet, i, type = "row") {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  assert_workbook(wb)

  sheet <- wb_validate_sheet(wb, sheet)

  type <- tolower(type)[1]
  if (!type %in% c("row", "column")) {
    stop("'type' argument must be 'row' or 'column'.")
  }

  if (!is.numeric(i)) {
    stop("'i' must be numeric.")
  }
  i <- round(i)

  if (type == "row") {
    wb$worksheets[[sheet]]$rowBreaks <- c(
      wb$worksheets[[sheet]]$rowBreaks,
      sprintf('<brk id="%s" max="16383" man="1"/>', i)
    )
  } else if (type == "column") {
    wb$worksheets[[sheet]]$colBreaks <- c(
      wb$worksheets[[sheet]]$colBreaks,
      sprintf('<brk id="%s" max="1048575" man="1"/>', i)
    )
  }


  # wb$worksheets[[sheet]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', paste(getCellRefs(data.frame("x" = c(rows, rows), "y" = c(min(cols), max(cols)))), collapse = ":"))

  invisible(wb)
}


#' @name getTables
#' @title List Excel tables in a workbook
#' @description List Excel tables in a workbook
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @return character vector of table names on the specified sheet
#' @examples
#'
#' wb <- wb_workbook()
#' wb$addWorksheet(sheet = "Sheet 1")
#' writeDataTable(wb, sheet = "Sheet 1", x = iris)
#' writeDataTable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#'
#' getTables(wb, sheet = "Sheet 1")
#' @export
getTables <- function(wb, sheet) {
  assert_workbook(wb)

  if (length(sheet) != 1) {
    stop("sheet argument must be length 1")
  }

  if (length(wb$tables) == 0) {
    return(character())
  }

  sheet <- wb_validate_sheet(wb, sheet)

  table_sheets <- attr(wb$tables, "sheet")
  tables <- attr(wb$tables, "tableName")
  refs <- names(wb$tables)

  refs <- refs[table_sheets == sheet & !grepl("openxlsx_deleted", tables, fixed = TRUE)]
  tables <- tables[table_sheets == sheet & !grepl("openxlsx_deleted", tables, fixed = TRUE)]

  if (length(tables)) {
    attr(tables, "refs") <- refs
  }

  return(tables)
}



#' @name removeTable
#' @title Remove an Excel table in a workbook
#' @description List Excel tables in a workbook
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param table Name of table to remove. See [getTables()]
#' @return character vector of table names on the specified sheet
#' @examples
#'
#' wb <- wb_workbook()
#' wb$addWorksheet(sheet = "Sheet 1")
#' wb$addWorksheet(sheet = "Sheet 2")
#' writeDataTable(wb, sheet = "Sheet 1", x = iris, tableName = "iris")
#' writeDataTable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#'
#'
#' wb <- wb_remove_worksheet(wb, sheet = 1) ## delete worksheet removes table objects
#'
#' writeDataTable(wb, sheet = 1, x = iris, tableName = "iris")
#' writeDataTable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#'
#' ## removeTable() deletes table object and all data
#' getTables(wb, sheet = 1)
#' removeTable(wb = wb, sheet = 1, table = "iris")
#' writeDataTable(wb, sheet = 1, x = iris, tableName = "iris", startCol = 1)
#'
#' getTables(wb, sheet = 1)
#' removeTable(wb = wb, sheet = 1, table = "iris")
#' writeDataTable(wb, sheet = 1, x = iris, tableName = "iris", startCol = 1)
#' \dontrun{
#' wb_save(wb, path = "removeTableExample.xlsx", overwrite = TRUE)
#' }
#'
#' @export
removeTable <- function(wb, sheet, table) {
  assert_workbook(wb)

  if (length(sheet) != 1) {
    stop("sheet argument must be length 1")
  }

  if (length(table) != 1) {
    stop("table argument must be length 1")
  }

  ## delete table object and all data in it
  sheet <- wb_validate_sheet(wb, sheet)

  if (!table %in% attr(wb$tables, "tableName")) {
    stop(sprintf("table '%s' does not exist.", table), call. = FALSE)
  }

  ## get existing tables
  table_sheets <- attr(wb$tables, "sheet")
  table_names <- attr(wb$tables, "tableName")
  refs <- names(wb$tables)

  ## delete table object (by flagging as deleted)
  inds <- which(table_sheets %in% sheet & table_names %in% table)
  table_name_original <- table_names[inds]

  table_names[inds] <- paste0(table_name_original, "_openxlsx_deleted")
  attr(wb$tables, "tableName") <- table_names

  ## delete reference from worksheet to table
  worksheet_table_names <- attr(wb$worksheets[[sheet]]$tableParts, "tableName")
  to_remove <- which(worksheet_table_names == table_name_original)

  wb$worksheets[[sheet]]$tableParts <- wb$worksheets[[sheet]]$tableParts[-to_remove]
  attr(wb$worksheets[[sheet]]$tableParts, "tableName") <- worksheet_table_names[-to_remove]


  ## Now delete data from the worksheet
  refs <- strsplit(refs[[inds]], split = ":")[[1]]
  rows <- as.integer(gsub("[A-Z]", "", refs))
  rows <- seq(from = rows[1], to = rows[2], by = 1)

  cols <- col2int(refs)
  cols <- seq(from = cols[1], to = cols[2], by = 1)

  ## now delete data
  deleteData(wb = wb, sheet = sheet, rows = rows, cols = cols, gridExpand = TRUE)

  invisible(0)
}


# grouping ----------------------------------------------------------------

#' Group Rows and Columns
#'
#' Group a selection of rows or cols
#'
#' @details If row was previously hidden, it will now be shown
#'
#' @param wb A [wbWorkbook] object
#' @param sheet A name or index of a worksheet
#' @param rows,cols Indices of rows and columns to group
#' @param collapsed If `TRUE` the grouped columns are collapsed
#' @param levels levels
#'
#' @examples
#' # create matrix
#' t1 <- AirPassengers
#' t2 <- do.call(cbind, split(t1, cycle(t1)))
#' dimnames(t2) <- dimnames(.preformat.ts(t1))
#'
#' wb <- wb_workbook()
#' wb$addWorksheet("AirPass")
#' writeData(wb, "AirPass", t2, rowNames = TRUE)
#'
#' # groups will always end on/show the last row. in the example 1950, 1955, and 1960
#' wb <- wb_group_rows(wb, "AirPass", 2:3, collapsed = TRUE) # group years < 1950
#' wb <- wb_group_rows(wb, "AirPass", 4:8, collapsed = TRUE) # group years 1951-1955
#' wb <- wb_group_rows(wb, "AirPass", 9:13)                  # group years 1956-1960
#'
#' wb$createCols("AirPass", 13)
#'
#' wb <- wb_group_cols(wb, "AirPass", 2:4, collapsed = TRUE)
#' wb <- wb_group_cols(wb, "AirPass", 5:7, collapsed = TRUE)
#' wb <- wb_group_cols(wb, "AirPass", 8:10, collapsed = TRUE)
#' wb <- wb_group_cols(wb, "AirPass", 11:13)
#'
#' @name workbook_grouping
#' @family workbook wrappers
NULL

#' @export
#' @rdname workbook_grouping
wb_group_cols <- function(wb, sheet, cols, collapsed = FALSE, levels = NULL) {
  assert_workbook(wb)
  wb$clone()$groupCols(
    sheet     = sheet,
    cols      = cols,
    collapsed = collapsed,
    levels    = levels
  )
}

#' @export
#' @rdname workbook_grouping
ungroupColumns <- function(wb, sheet, cols) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  sheet <- wb_validate_sheet(wb, sheet)

  # check if any rows are selected
  if (any(cols) < 1L) {
    stop("Invalid cols entered (<= 0).")
  }

  # fetch the cols_attr data.frame
  col_attr <- wb$worksheets[[sheet]]$unfold_cols()

  # get the selection based on the col_attr frame.
  select <- col_attr$min %in% as.character(cols)
  if (length(select)) {
    col_attr$outlineLevel[select] <- ""
    col_attr$collapsed[select] <- ""
    # TODO only if unhide = TRUE
    col_attr$hidden[select] <- ""
    wb$worksheets[[sheet]]$fold_cols(col_attr)
  }

  # If all outlineLevels are missing: remove the outlineLevelCol attribute. Assigning "" will remove the attribute
  if (all(col_attr$outlineLevel == "")) {
    wb$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(wb$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelCol = ""))
  } else {
    self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelCol = as.character(max(as.integer(col_attr$outlineLevel)))))
  }
}


#' @export
#' @rdname workbook_grouping
wb_group_rows <- function(wb, sheet, rows, collapsed = FALSE, levels = NULL) {
  assert_workbook(wb)
  wb$clone()$groupRows(
    sheet     = sheet,
    rows      = rows,
    collapsed = collapsed,
    levels    = levels
  )
}

#' @export
#' @rdname workbook_grouping
ungroupRows <- function(wb, sheet, rows) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  sheet <- wb_validate_sheet(wb, sheet)

  # check if any rows are selected
  if (any(rows) < 1L) {
    stop("Invalid rows entered (<= 0).")
  }

  # fetch the row_attr data.frame
  row_attr <- wb$worksheets[[sheet]]$sheet_data$row_attr

  # get the selection based on the row_attr frame.
  select <- row_attr$r %in% as.character(rows)
  if (length(select)) {
    row_attr$outlineLevel[select] <- ""
    row_attr$collapsed[select] <- ""
    # TODO only if unhide = TRUE
    row_attr$hidden[select] <- ""
    wb$worksheets[[sheet]]$sheet_data$row_attr <- row_attr
  }

  # If all outlineLevels are missing: remove the outlineLevelRow attribute. Assigning "" will remove the attribute
  if (all(row_attr$outlineLevel == "")) {
    wb$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(wb$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelRow = ""))
  } else {
    self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelRow = as.character(max(as.integer(row_attr$outlineLevel)))))
  }
}


# creators ----------------------------------------------------------------

#' Workbook creators
#'
#' Modify and get workbook creators
#'
#' @param wb A `wbWorkbook` object
#' @examples
#'
#' # workbook made with default creator (see [wbWorkbook])
#' wb <- wb_workbook()
#' wb_get_creators(wb)
#'
#' # add a new creator (assuming "test" isn't your default creator)
#' wb <- wb_add_creators(wb, "test")
#' wb_get_creators(wb)
#'
#' # remove the creator (should be the same as before)
#' wb <- wb_remove_creators(wb, "test")
#' wb_get_creators(wb)
#'
#' @return
#' * `wb_set_creators()`, `wb_add_creators()`, and `wb_remove_creators()` return the `wbWorkbook` object
#' * `wb_get_creators()` returns a `character` vector of creators
#'
#' @name wb_creators
#' @family workbook wrappers
NULL

#' @rdname wb_creators
#' @export
#' @param creators A character vector of names
wb_add_creators <- function(wb, creators) {
  assert_workbook(wb)
  wb$clone()$addCreators(creators)
}

#' @rdname wb_creators
#' @export
wb_set_creators <- function(wb, creators) {
  assert_workbook(wb)
  wb$clone()$setCreators(creators)
}

#' @rdname wb_creators
#' @export
wb_remove_creators <- function(wb, creators) {
  assert_workbook(wb)
  wb$clone()$removeCreators(creators)
}

#' @rdname wb_creators
#' @export
wb_get_creators <- function(wb) {
  assert_workbook(wb)
  wb[["creator"]]
}


# others? -----------------------------------------------------------------

#' Add another author to the meta data of the file.
#'
#' Just a wrapper of wb$changeLastModifiedBy()
#'
#' @param wb A workbook object
#' @param LastModifiedBy A string object with the name of the LastModifiedBy-User
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' wb <- wb_workbook()
#' setLastModifiedBy(wb, "test")
setLastModifiedBy <- function(wb, LastModifiedBy) {
  assert_workbook(wb)
  wb$changeLastModifiedBy(LastModifiedBy)
}

#' @name insertImage
#' @title Insert an image into a worksheet
#' @description Insert an image into a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param file An image file. Valid file types are: jpeg, png, bmp
#' @param width Width of figure.
#' @param height Height of figure.
#' @param startRow Row coordinate of upper left corner of the image
#' @param startCol Column coordinate of upper left corner of the image
#' @param units Units of width and height. Can be "in", "cm" or "px"
#' @param dpi Image resolution used for conversion between units.
#' @seealso [insertPlot()]
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook("Ayanami")
#'
#' ## Add some worksheets
#' wb$addWorksheet("Sheet 1")
#' wb$addWorksheet("Sheet 2")
#' wb$addWorksheet("Sheet 3")
#'
#' ## Insert images
#' img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
#' insertImage(wb, "Sheet 1", img, startRow = 5, startCol = 3, width = 6, height = 5)
#' insertImage(wb, 2, img, startRow = 2, startCol = 2)
#' insertImage(wb, 3, img, width = 15, height = 12, startRow = 3, startCol = "G", units = "cm")
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "insertImageExample.xlsx", overwrite = TRUE)
#' }
insertImage <- function(wb, sheet, file, width = 6, height = 3, startRow = 1, startCol = 1, units = "in", dpi = 300) {
  op <- openxlsx_options()
  on.exit(options(op), add = TRUE)

  if (!file.exists(file)) {
    stop("File does not exist.")
  }

  if (!grepl("\\\\|\\/", file)) {
    file <- file.path(getwd(), file, fsep = .Platform$file.sep)
  }

  units <- tolower(units)

  if (!units %in% c("cm", "in", "px")) {
    stop("Invalid units.\nunits must be one of: cm, in, px")
  }

  startCol <- col2int(startCol)
  startRow <- as.integer(startRow)

  ## convert to inches
  if (units == "px") {
    width <- width / dpi
    height <- height / dpi
  } else if (units == "cm") {
    width <- width / 2.54
    height <- height / 2.54
  }

  ## Convert to EMUs
  widthEMU <- as.integer(round(width * 914400L, 0)) # (EMUs per inch)
  heightEMU <- as.integer(round(height * 914400L, 0)) # (EMUs per inch)

  wb$insertImage(sheet, file = file, startRow = startRow, startCol = startCol, width = widthEMU, height = heightEMU)
}
