
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
#' wb$add_worksheet(sheet = "My first worksheet")
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
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
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
#' wb <- wb_merge_cells(wb, 2, cols = c(1, 10), rows = 12) # equivalent to 1:10
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
  wb$clone()$merge_cells(sheet, rows = rows, cols = cols)
}

#' @export
#' @rdname ws_cell_merge
wb_unmerge_cells <- function(wb, sheet, rows = NULL, cols = NULL) {
  assert_workbook(wb)
  wb$clone()$unmerge_cells(sheet, rows = rows, cols = cols)
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
#' @param paperSize An integer corresponding to a paper size. See ?ws_page_setup for
#'   details.
#' @param orientation One of "portrait" or "landscape"
#' @param hdpi Horizontal DPI. Can be set with options("openxlsx2.dpi" = X) or
#'   options("openxlsx2.hdpi" = X)
#' @param vdpi Vertical DPI. Can be set with options("openxlsx2.dpi" = X) or
#'   options("openxlsx2.vdpi" = X)
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
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2", gridLines = FALSE)
#' wb$add_worksheet("Sheet 3", tabColour = "red")
#' wb$add_worksheet("Sheet 4", gridLines = FALSE, tabColour = "#4F81BD")
#'
#' ## Headers and Footers
#' wb$add_worksheet("Sheet 5",
#'   header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
#'   footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
#'   evenHeader = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
#'   evenFooter = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
#'   firstHeader = c("TOP", "OF FIRST", "PAGE"),
#'   firstFooter = c("BOTTOM", "OF FIRST", "PAGE")
#' )
#'
#' wb$add_worksheet("Sheet 6",
#'   header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
#'   footer = c("&[Path]&[File]", NA, "&[Tab]"),
#'   firstHeader = c(NA, "Center Header of First Page", NA),
#'   firstFooter = c(NA, "Center Footer of First Page", NA)
#' )
#'
#' wb$add_worksheet("Sheet 7",
#'   header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
#'   footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2")
#' )
#'
#' wb$add_worksheet("Sheet 8",
#'   firstHeader = c("FIRST ONLY L", NA, "FIRST ONLY R"),
#'   firstFooter = c("FIRST ONLY L", NA, "FIRST ONLY R")
#' )
#'
#' ## Need data on worksheet to see all headers and footers
#' write_data(wb, sheet = 5, 1:400)
#' write_data(wb, sheet = 6, 1:400)
#' write_data(wb, sheet = 7, 1:400)
#' write_data(wb, sheet = 8, 1:400)
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "add_worksheetExample.xlsx", overwrite = TRUE)
#' }
wb_add_worksheet <- function(
  wb,
  sheet,
  gridLines   = TRUE,
  tabColour   = NULL,
  zoom        = 100,
  header      = NULL,
  footer      = NULL,
  oddHeader   = header,
  oddFooter   = footer,
  evenHeader  = header,
  evenFooter  = footer,
  firstHeader = header,
  firstFooter = footer,
  visible     = c("true", "false", "hidden", "visible", "veryhidden"),
  hasDrawing  = FALSE,
  paperSize   = getOption("openxlsx2.paperSize", default = 9),
  orientation = getOption("openxlsx2.orientation", default = "portrait"),
  hdpi        = getOption("openxlsx2.hdpi", default = getOption("openxlsx2.dpi", default = 300)),
  vdpi        = getOption("openxlsx2.vdpi", default = getOption("openxlsx2.dpi", default = 300))
) {
  assert_workbook(wb)
  wb$clone()$add_worksheet(
    sheet       = sheet,
    gridLines   = gridLines,
    tabColour   = tabColour,
    zoom        = zoom,
    oddHeader   = headerFooterSub(oddHeader),
    oddFooter   = headerFooterSub(oddFooter),
    evenHeader  = headerFooterSub(evenHeader),
    evenFooter  = headerFooterSub(evenFooter),
    firstHeader = headerFooterSub(firstHeader),
    firstFooter = headerFooterSub(firstFooter),
    visible     = visible,
    paperSize   = paperSize,
    orientation = orientation,
    vdpi        = vdpi,
    hdpi        = hdpi
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
#' wb$add_worksheet("Sheet 1")
#' wb$clone_worksheet("Sheet 1", "Sheet 2")
#'
#' # Save workbook
#' \dontrun{
#' wb_save(wb, "clone_worksheetExample.xlsx", overwrite = TRUE)
#' }
wb_clone_worksheet <- function(wb, old, new) {
  assert_workbook(wb)
  wb$clone()$clone_worksheet(old = old, new = new)
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
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#' wb$add_worksheet("Sheet 3")
#' wb$add_worksheet("Sheet 4")
#'
#' ## Freeze Panes
#' wb$freeze_pane("Sheet 1", firstActiveRow = 5, firstActiveCol = 3)
#' wb$freeze_pane("Sheet 2", firstCol = TRUE) ## shortcut to firstActiveCol = 2
#' wb$freeze_pane(3, firstRow = TRUE) ## shortcut to firstActiveRow = 2
#' wb$freeze_pane(4, firstActiveRow = 1, firstActiveCol = "D")
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "freezePaneExample.xlsx", overwrite = TRUE)
#' }
wb_freeze_pane <- function(wb, sheet, firstActiveRow = NULL, firstActiveCol = NULL, firstRow = FALSE, firstCol = FALSE) {
  assert_workbook(wb)
  wb$clone()$freeze_pane(
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
#' @seealso [wb_remove_row_heights()]
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$add_worksheet("Sheet 1")
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
#' wb_save(wb, "set_row_heightsExample.xlsx", overwrite = TRUE)
#' }
wb_set_row_heights <- function(wb, sheet, rows, heights) {
  assert_workbook(wb)
  wb$clone()$set_row_heights(sheet, rows, heights)
}


#' Set worksheet column widths
#'
#' Set worksheet column widths to specific width or "auto".
#'
#' @param wb A [wbWorkbook] object
#' @param sheet A name or index of a worksheet
#' @param cols Indices of cols to set width
#' @param widths width to set cols to specified in Excel column width units or "auto" for automatic sizing. The widths argument is
#' recycled to the length of cols. The default width is 8.43. Though there is no specific default width for Excel, it depends on
#' Excel version, operating system and DPI settings used. Setting it to specific value also is no guarantee that the output will be
#' of the selected width.
#' @param hidden Logical vector. If TRUE the column is hidden.
#' @details The global min and max column width for "auto" columns is set by (default values show):
#' \itemize{
#'   \item{options("openxlsx2.minWidth" = 3)}
#'   \item{options("openxlsx2.maxWidth" = 250)} ## This is the maximum width allowed in Excel
#' }
#'
#'   NOTE: The calculation of column widths can be slow for large worksheets.
#'
#'   NOTE: The `hidden` parameter may conflict with the one set in
#'   [wb_group_cols]; changing one will update the other.
#'
#' @export
#' @family workbook wrappers
#' @seealso [wb_remove_col_widths()]
#'
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$add_worksheet("Sheet 1")
#'
#' ## set col widths
#' wb$set_col_widths(1, cols = c(1, 4, 6, 7, 9), widths = c(16, 15, 12, 18, 33))
#'
#' ## auto columns
#' wb$add_worksheet("Sheet 2")
#' write_data(wb, sheet = 2, x = iris)
#' wb$set_col_widths(sheet = 2, cols = 1:5, widths = "auto")
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "wb_set_col_widthsExample.xlsx", overwrite = TRUE)
#' }
#'
wb_set_col_widths <- function(wb, sheet, cols, widths = 8.43, hidden = FALSE) {
  assert_workbook(wb)
  wb$clone()$set_col_widths(
    sheet  = sheet,
    cols   = cols,
    widths = widths,
    # TODO allow either 1 or length(cols)
    hidden = hidden
  )
}

#' @name wb_remove_col_widths
#' @title Remove column widths from a worksheet

#' @description Remove column widths from a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Indices of columns to remove custom width (if any) from.
#' @seealso [wb_set_col_widths()]
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- wb_load(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#'
#' ## remove column widths in columns 1 to 20
#' wb_remove_col_widths(wb, 1, cols = 1:20)
#' \dontrun{
#' wb_save(wb, "wb_remove_col_widthsExample.xlsx", overwrite = TRUE)
#' }
wb_remove_col_widths <- function(wb, sheet, cols) {
  assert_workbook(wb)
  wb$clone()$remove_col_widths(sheet = sheet, cols = cols)
}



#' Remove custom row heights from a worksheet
#'
#' Remove row heights from a worksheet
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Indices of rows to remove custom height (if any) from.
#' @seealso [wb_set_row_heights()]
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- wb_load(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#'
#' ## remove any custom row heights in rows 1 to 10
#' wb$remove_row_heights(1, rows = 1:10)
#' \dontrun{
#' wb_save(wb, "remove_row_heightsExample.xlsx", overwrite = TRUE)
#' }
wb_remove_row_heights <- function(wb, sheet, rows) {
  assert_workbook(wb)
  wb$clone()$remove_row_heights(sheet = sheet, rows = rows)
}


# images ------------------------------------------------------------------


#' Insert the current plot into a worksheet
#'
#' The current plot is saved to a temporary image file using
#' [grDevices::dev.copy()] This file is then written to the workbook using
#' [wb_add_image()].
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param xy Alternate way to specify `startRow` and `startCol.`  A vector of
#'   length `2` of form (`startcol`, `startRow`)
#' @param startRow Row coordinate of upper left corner of figure. `xy[[2]]` when
#'   `xy` is given.
#' @param startCol Column coordinate of upper left corner of figure. `xy[[1]]`
#'   when `xy` is given.
#' @param rowOffset offset within cell (row)
#' @param colOffset offset within cell (column)
#' @param width Width of figure. Defaults to `6`in.
#' @param height Height of figure . Defaults to `4`in.
#' @param fileType File type of image
#' @param units Units of width and height. Can be `"in"`, `"cm"` or `"px"`
#' @param dpi Image resolution
#' @seealso [wb_add_image()]
#' @export
#' @examples
#' \dontrun{
#' ## Create a new workbook
#' wb <- wb_workbook()
#'
#' ## Add a worksheet
#' wb$add_worksheet("Sheet 1", gridLines = FALSE)
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
#' wb$add_plot(1, width = 5, height = 3.5, fileType = "png", units = "in")
#'
#' ## Insert plot 2
#' print(p2)
#' wb$add_plot(1, xy = c("J", 2), width = 16, height = 10, fileType = "png", units = "cm")
#'
#' ## Save workbook
#' wb_save(wb, "wb_add_plotExample.xlsx", overwrite = TRUE)
#' }
wb_add_plot <- function(
  wb,
  sheet,
  width     = 6,
  height    = 4,
  xy        = NULL,
  startRow  = 1,
  startCol  = 1,
  rowOffset = 0,
  colOffset = 0,
  fileType  = "png",
  units     = "in",
  dpi       = 300
) {
  assert_workbook(wb)
  wb$clone()$add_plot(
    sheet     = sheet,
    width     = width,
    height    = height,
    xy        = xy,
    startRow  = startRow,
    startCol  = startCol,
    rowOffset = rowOffset,
    colOffset = colOffset,
    fileType  = fileType,
    units     = units,
    dpi       = dpi
  )
}


#' @title Remove a worksheet from a workbook
#' @description Remove a worksheet from a Workbook object
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @description Remove a worksheet from a workbook
#' @export
#' @examples
#' ## load a workbook
#' wb <- wb_load(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#'
#' ## Remove sheet 2
#' wb <- wb_remove_worksheet(wb, 2)
#'
#' ## save the modified workbook
#' \dontrun{
#' wb_save(wb, "remove_worksheetExample.xlsx", overwrite = TRUE)
#' }
wb_remove_worksheet <- function(wb, sheet) {
  assert_workbook(wb)
  wb$clone()$remove_worksheet(sheet)
}


# base font ---------------------------------------------------------------

#' @name wb_modify_basefont
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
#' wb$add_worksheet("S1")
#' ## modify base font to size 10 Arial Narrow in red
#' wb$set_base_font(fontSize = 10, fontColour = "#FF0000", fontName = "Arial Narrow")
#'
#' write_data(wb, "S1", iris)
#' write_datatable(wb, "S1", x = iris, startCol = 10) ## font colour does not affect tables
#' \dontrun{
#' wb_save(wb, "wb_set_base_font_example.xlsx", overwrite = TRUE)
#' }
wb_set_base_font <- function(wb, fontSize = 11, fontColour = "black", fontName = "Calibri") {
  assert_workbook(wb)
  wb$clone()$set_base_font(
    fontSize   = fontSize,
    fontColour = fontColour,
    fontName   = fontName
  )
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
#' wb$set_base_font(fontSize = 10, fontColour = "#FF0000", fontName = "Arial Narrow")
#'
#' wb_get_base_font(wb)
wb_get_base_font <- function(wb) {
  # TODO all of these class checks need to be cleaned up
  assert_workbook(wb)
  wb$get_base_font()
}

#' Set document headers and footers
#'
#' Set document headers and footers
#'
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
#' wb$add_worksheet("S1")
#' wb$add_worksheet("S2")
#' wb$add_worksheet("S3")
#' wb$add_worksheet("S4")
#'
#' write_data(wb, 1, 1:400)
#' write_data(wb, 2, 1:400)
#' write_data(wb, 3, 3:400)
#' write_data(wb, 4, 3:400)
#'
#' wb$set_header_footer(
#'   sheet = "S1",
#'   header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
#'   footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
#'   evenHeader = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
#'   evenFooter = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
#'   firstHeader = c("TOP", "OF FIRST", "PAGE"),
#'   firstFooter = c("BOTTOM", "OF FIRST", "PAGE")
#' )
#'
#' wb$set_header_footer(
#'   sheet = 2,
#'   header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
#'   footer = c("&[Path]&[File]", NA, "&[Tab]"),
#'   firstHeader = c(NA, "Center Header of First Page", NA),
#'   firstFooter = c(NA, "Center Footer of First Page", NA)
#' )
#'
#' wb$set_header_footer(
#'   sheet = 3,
#'   header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
#'   footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2")
#' )
#'
#' wb$set_header_footer(
#'   sheet = 4,
#'   firstHeader = c("FIRST ONLY L", NA, "FIRST ONLY R"),
#'   firstFooter = c("FIRST ONLY L", NA, "FIRST ONLY R")
#' )
#' \dontrun{
#' wb_save(wb, "wb_set_header_footerExample.xlsx", overwrite = TRUE)
#' }
wb_set_header_footer <- function(
  wb,
  sheet,
  header = NULL,
  footer = NULL,
  evenHeader = NULL,
  evenFooter = NULL,
  firstHeader = NULL,
  firstFooter = NULL
) {
  assert_workbook(wb)
  wb$clone()$set_header_footer(
    sheet       = sheet,
    header      = header,
    footer      = footer,
    evenHeader  = evenHeader,
    evenFooter  = evenFooter,
    firstHeader = firstHeader,
    firstFooter = firstFooter
  )
}



#' @name ws_page_setup
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
#' wb$add_worksheet("S1")
#' wb$add_worksheet("S2")
#' write_datatable(wb, 1, x = iris[1:30, ])
#' write_datatable(wb, 2, x = iris[1:30, ], xy = c("C", 5))
#'
#' ## landscape page scaled to 50%
#' wb$page_setup(sheet = 1, orientation = "landscape", scale = 50)
#'
#' ## portrait page scales to 300% with 0.5in left and right margins
#' wb$page_setup(sheet = 2, orientation = "portrait", scale = 300, left = 0.5, right = 0.5)
#'
#'
#' ## print titles
#' wb$add_worksheet("print_title_rows")
#' wb$add_worksheet("print_title_cols")
#'
#' write_data(wb, "print_title_rows", rbind(iris, iris, iris, iris))
#' write_data(wb, "print_title_cols", x = rbind(mtcars, mtcars, mtcars), rowNames = TRUE)
#'
#' wb$page_setup(sheet = "print_title_rows", printTitleRows = 1) ## first row
#' wb$page_setup(sheet = "print_title_cols", printTitleCols = 1, printTitleRows = 1)
#' \dontrun{
#' wb_save(wb, "ws_page_setupExample.xlsx", overwrite = TRUE)
#' }
wb_page_setup <- function(
  wb,
  sheet,
  orientation    = NULL,
  scale          = 100,
  left           = 0.7,
  right          = 0.7,
  top            = 0.75,
  bottom         = 0.75,
  header         = 0.3,
  footer         = 0.3,
  fitToWidth     = FALSE,
  fitToHeight    = FALSE,
  paperSize      = NULL,
  printTitleRows = NULL,
  printTitleCols = NULL,
  summaryRow     = NULL,
  summaryCol     = NULL
) {
  assert_workbook(wb)
  wb$clone()$page_setup(
    sheet          = sheet,
    orientation    = orientation,
    scale          = scale,
    left           = left,
    right          = right,
    top            = top,
    bottom         = bottom,
    header         = header,
    footer         = footer,
    fitToWidth     = fitToWidth,
    fitToHeight    = fitToHeight,
    paperSize      = paperSize,
    printTitleRows = printTitleRows,
    printTitleCols = printTitleCols,
    summaryRow     = summaryRow,
    summaryCol     = summaryCol
  )
}


# protect -----------------------------------------------------------------

#' @name ws_protect
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
#' wb$add_worksheet("S1")
#' write_datatable(wb, 1, x = iris[1:30, ])
#' # Formatting cells / columns is allowed , but inserting / deleting columns is protected:
#' ws_protect(wb, "S1",
#'   protect = TRUE,
#'   lockFormattingCells = FALSE, lockFormattingColumns = FALSE,
#'   lockInsertingColumns = TRUE, lockDeletingColumns = TRUE
#' )
#'
#' # Remove the protection
#' ws_protect(wb, "S1", protect = FALSE)
#' \dontrun{
#' wb_save(wb, "ws_page_setupExample.xlsx", overwrite = TRUE)
#' }
ws_protect <- function(wb, sheet, protect = TRUE, password = NULL,
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


#' @name wb_protect
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
#' wb$add_worksheet("S1")
#' wb_protect(wb, protect = TRUE, password = "Password", lockStructure = TRUE)
#' \dontrun{
#' wb_save(wb, "WorkBook_Protection.xlsx", overwrite = TRUE)
#' }
#' # Remove the protection
#' wb_protect(wb, protect = FALSE)
#' \dontrun{
#' wb_save(wb, "WorkBook_Protection_unprotected.xlsx", overwrite = TRUE)
#' }
#'
#' wb_protect(
#'   wb,
#'   protect = TRUE,
#'   password = "Password",
#'   lockStructure = TRUE,
#'   type = 2L,
#'   fileSharing = TRUE,
#'   username = "Test",
#'   readOnlyRecommended = TRUE
#' )
#'
wb_protect <- function(wb, protect = TRUE, password = NULL, lockStructure = FALSE, lockWindows = FALSE, type = 1L, fileSharing = FALSE, username = unname(Sys.info()["user"]), readOnlyRecommended = FALSE) {
  assert_workbook(wb)
  invisible(wb$protect(protect = protect, password = password, lockStructure = lockStructure, lockWindows = lockWindows, type = type, fileSharing = fileSharing, username = username, readOnlyRecommended = readOnlyRecommended))
}


# grid lines --------------------------------------------------------------

#' Set worksheet gridlines to show or hide.
#'
#' Set worksheet gridlines to show or hide.
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param show A logical. If `FALSE`, grid lines are hidden.
#' @export
#' @examples
#' wb <- wb_load(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#' names(wb) ## list worksheets in workbook
#' wb$grid_lines(1, show = FALSE)
#' wb$grid_lines("testing", show = FALSE)
#' \dontrun{
#' wb_save(wb, "ws_grid_linesExample.xlsx", overwrite = TRUE)
#' }
wb_grid_lines <- function(wb, sheet, show = FALSE) {
  assert_workbook(wb)
  wb$clone()$grid_lines(sheet = sheet, show = show)
}

# TODO hide gridlines?

# worksheet order ---------------------------------------------------------

#' Order of worksheets in xlsx file
#'
#' Get/set order of worksheets in a Workbook object
#'
#' @param wb A `wbWorkbook` object
#'
#' @details This function does not reorder the worksheets within the workbook
#'   object, it simply shuffles the order when writing to file.
#' @export
#' @examples
#' ## setup a workbook with 3 worksheets
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1", gridLines = FALSE)
#' write_datatable(wb = wb, sheet = 1, x = iris)
#'
#' wb$add_worksheet("mtcars (Sheet 2)", gridLines = FALSE)
#' write_data(wb = wb, sheet = 2, x = mtcars)
#'
#' wb$add_worksheet("Sheet 3", gridLines = FALSE)
#' write_data(wb = wb, sheet = 3, x = Formaldehyde)
#'
#' wb_get_order(wb)
#' names(wb)
#' wb$set_order(c(1, 3, 2)) # switch position of sheets 2 & 3
#' write_data(wb, 2, 'This is still the "mtcars" worksheet', startCol = 15)
#' wb_get_order(wb)
#' names(wb) ## ordering within workbook is not changed
#' \dontrun{
#' wb_save(wb, "wb_orderExample.xlsx", overwrite = TRUE)
#' }
#' wb$set_order(3:1)
#' \dontrun{
#' wb_save(wb, "wb_orderExample2.xlsx", overwrite = TRUE)
#' }
#' @name wb_order
wb_get_order <- function(wb) {
  assert_workbook(wb)
  wb$sheetOrder
}

#' @rdname wb_order
#' @param sheets Sheet order
#' @export
wb_set_order <- function(wb, sheets) {
  assert_workbook(wb)
  wb$clone()$set_order(sheets = sheets)
}


# named region ------------------------------------------------------------


#' Create / delete a named region
#'
#' Create / delete a named region
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Numeric vector specifying rows to include in region
#' @param cols Numeric vector specifying columns to include in region
#' @param name Name for region. A character vector of length 1. Note region names musts be case-insensitive unique.
#' @param overwrite Boolean. Overwrite if exists? Default to FALSE
#' @param localSheetId localSheetId
#' @details Region is given by: min(cols):max(cols) X min(rows):max(rows)
#' @examples
#' ## create named regions
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#'
#' ## specify region
#' write_data(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
#' wb$add_named_region(
#'   sheet = 1,
#'   name = "iris",
#'   rows = seq_len(nrow(iris) + 1),
#'   cols = seq_along(iris)
#' )
#'
#'
#' ## using write_data 'name' argument
#' write_data(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
#'
#' out_file <- tempfile(fileext = ".xlsx")
#' \dontrun{
#' wb_save(wb, out_file, overwrite = TRUE)
#'
#' ## see named regions
#' get_named_regions(wb) ## From Workbook object
#' get_named_regions(out_file) ## From xlsx file
#'
#' ## delete one
#' wb$remove_named_region(name = "iris2")
#' get_named_regions(wb)
#'
#' ## read named regions
#' df <- read_xlsx(wb, namedRegion = "iris")
#' head(df)
#'
#' df <- read_xlsx(out_file, namedRegion = "iris2")
#' head(df)
#' }
#' @name named_region
NULL

#' @rdname named_region
#' @export
wb_add_named_region <- function(wb, sheet, cols, rows, name, localSheetId = NULL, overwrite = FALSE) {
  assert_workbook(wb)
  wb$clone()$add_named_region(
    sheet        = sheet,
    cols         = cols,
    rows         = rows,
    name         = name,
    localSheetId = localSheetId,
    overwrite    = overwrite
  )
}

#' @rdname named_region
#' @export
wb_remove_named_region <- function(wb, sheet = NULL, name = NULL) {
  assert_workbook(wb)
  wb$clone()$remove_named_region(sheet = sheet, name = name)
}

# filters -----------------------------------------------------------------

#' Add column filters
#'
#' Add excel column filters to a worksheet
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols columns to add filter to.
#' @param rows A row number.
#' @seealso [write_data()]
#' @details adds filters to worksheet columns, same as filter parameters in write_data.
#' write_datatable automatically adds filters to first row of a table.
#' NOTE Can only have a single filter per worksheet unless using tables.
#' @export
#' @seealso [wb_add_filter()]
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#' wb$add_worksheet("Sheet 3")
#'
#' write_data(wb, 1, iris)
#' wb$add_filter(1, row = 1, cols = seq_along(iris))
#'
#' ## Equivalently
#' write_data(wb, 2, x = iris, withFilter = TRUE)
#'
#' ## Similarly
#' write_datatable(wb, 3, iris)
#' \dontrun{
#' wb_save(wb, path = "wb_add_filterExample.xlsx", overwrite = TRUE)
#' }
wb_add_filter <- function(wb, sheet, rows, cols) {
  assert_workbook(wb)
  wb$clone()$add_filter(sheet = sheet, rows = rows, cols = cols)
}

#' @name wb_remove_filter
#' @title Remove a worksheet filter
#' @description Removes filters from wb_add_filter() and write_data()
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @export
#' @examples
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#' wb$add_worksheet("Sheet 3")
#'
#' write_data(wb, 1, iris)
#' wb_add_filter(wb, 1, row = 1, cols = seq_along(iris))
#'
#' ## Equivalently
#' write_data(wb, 2, x = iris, withFilter = TRUE)
#'
#' ## Similarly
#' write_datatable(wb, 3, iris)
#'
#' ## remove filters
#' wb_remove_filter(wb, 1:2) ## remove filters
#' wb_remove_filter(wb, 3) ## Does not affect tables!
#' \dontrun{
#' wb_save(wb, path = "wb_remove_filterExample.xlsx", overwrite = TRUE)
#' }
wb_remove_filter <- function(wb, sheet) {
  assert_workbook(wb)
  wb$clone()$remove_filter(sheet = sheet)
}


# validations -------------------------------------------------------------

#' Add data validation to cells
#'
#' Add Excel data validation to cells
#'
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
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#'
#' write_datatable(wb, 1, x = iris[1:30, ])
#' wb$add_data_validation(1,
#'   col = 1:3, rows = 2:31, type = "whole",
#'   operator = "between", value = c(1, 9)
#' )
#' wb$add_data_validation(1,
#'   col = 5, rows = 2:31, type = "textLength",
#'   operator = "between", value = c(4, 6)
#' )
#'
#' ## Date and Time cell validation
#' df <- data.frame(
#'   "d" = as.Date("2016-01-01") + -5:5,
#'   "t" = as.POSIXct("2016-01-01") + -5:5 * 10000
#' )
#' write_datatable(wb, 2, x = df)
#' wb$add_data_validation(2,
#'   col = 1, rows = 2:12, type = "date",
#'   operator = "greaterThanOrEqual", value = as.Date("2016-01-01")
#' )
#' wb$add_data_validation(2,
#'   col = 2, rows = 2:12, type = "time",
#'   operator = "between", value = df$t[c(4, 8)]
#' )
#'
#' \dontrun{
#' wb_save(wb, "data_validationExample.xlsx", overwrite = TRUE)
#' }
#'
#'
#' ######################################################################
#' ## If type == 'list'
#' # operator argument is ignored.
#'
#' wb <- wb_workbook()
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#'
#' write_datatable(wb, sheet = 1, x = iris[1:30, ])
#' write_data(wb, sheet = 2, x = sample(iris$Sepal.Length, 10))
#'
#' wb$add_data_validation(1, col = 1, rows = 2:31, type = "list", value = "'Sheet 2'!$A$1:$A$10")
#'
#' \dontrun{
#' wb_save(wb, "data_validationExample2.xlsx", overwrite = TRUE)
#' }
wb_add_data_validation <- function(
    wb,
    sheet,
    cols,
    rows,
    type,
    operator,
    value,
    allowBlank = TRUE,
    showInputMsg = TRUE,
    showErrorMsg = TRUE
) {
  assert_workbook(wb)
  wb$clone()$add_data_validation(
    sheet        = sheet,
    cols         = cols,
    rows         = rows,
    type         = type,
    operator     = operator,
    value        = value,
    allowBlank   = allowBlank,
    showInputMsg = showInputMsg,
    showErrorMsg = showErrorMsg
  )
}


# visibility --------------------------------------------------------------

#' Get/set worksheet visible state
#'
#'Get and set worksheet visible state
#'
#' @return Character vector of worksheet names.
#' @return  Vector of "hidden", "visible", "veryHidden"
#' @examples
#'
#' wb <- wb_workbook()
#' wb$add_worksheet(sheet = "S1", visible = FALSE)
#' wb$add_worksheet(sheet = "S2", visible = TRUE)
#' wb$add_worksheet(sheet = "S3", visible = FALSE)
#'
#' wb$get_sheet_visibility()
#' wb$set_sheet_visibility(1, TRUE)         ## show sheet 1
#' wb$set_sheet_visibility(2, FALSE)        ## hide sheet 2
#' wb$set_sheet_visibility(3, "hidden")     ## hide sheet 3
#' wb$set_sheet_visibility(3, "veryHidden") ## hide sheet 3 from UI
#' @name sheet_visibility
NULL

#' @rdname sheet_visibility
#' @param wb A `wbWorkbook` object
#' @export
wb_get_sheet_visibility <- function(wb) {
  assert_workbook(wb)
  wb$get_sheet_visibility()
}

#' @rdname sheet_visibility
#' @param sheet Worksheet identifier
#' @param value a logical/character vector the same length as sheet
#' @export
wb_set_sheet_visibility <- function(wb, sheet, value) {
  assert_workbook(wb)
  wb$clone()$set_sheet_visibility(sheet = sheet, value = value)
}


#' @name wb_page_break
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
#' wb$add_worksheet("Sheet 1")
#' write_data(wb, sheet = 1, x = iris)
#'
#' wb_page_break(wb, sheet = 1, i = 10, type = "row")
#' wb_page_break(wb, sheet = 1, i = 20, type = "row")
#' wb_page_break(wb, sheet = 1, i = 2, type = "column")
#' \dontrun{
#' wb_save(wb, "wb_page_breakExample.xlsx", TRUE)
#' }
#' ## In Excel: View tab -> Page Break Preview
wb_page_break <- function(wb, sheet, i, type = "row") {
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


  # wb$worksheets[[sheet]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', paste(get_cell_refs(data.frame("x" = c(rows, rows), "y" = c(min(cols), max(cols)))), collapse = ":"))

  invisible(wb)
}


#' @name wb_get_tables
#' @title List Excel tables in a workbook
#' @description List Excel tables in a workbook
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @return character vector of table names on the specified sheet
#' @examples
#'
#' wb <- wb_workbook()
#' wb$add_worksheet(sheet = "Sheet 1")
#' write_datatable(wb, sheet = "Sheet 1", x = iris)
#' write_datatable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#'
#' wb$get_tables(sheet = "Sheet 1")
#' @export
wb_get_tables <- function(wb, sheet) {
  assert_workbook(wb)
  wb$clone()$get_tables(sheet = sheet)
}



#' Remove an Excel table in a workbook
#'
#' List Excel tables in a workbook
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param table Name of table to remove. See [wb_get_tables()]
#' @return character vector of table names on the specified sheet
#' @examples
#'
#' wb <- wb_workbook()
#' wb$add_worksheet(sheet = "Sheet 1")
#' wb$add_worksheet(sheet = "Sheet 2")
#' write_datatable(wb, sheet = "Sheet 1", x = iris, tableName = "iris")
#' write_datatable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#'
#'
#' wb <- wb_remove_worksheet(wb, sheet = 1) ## delete worksheet removes table objects
#'
#' write_datatable(wb, sheet = 1, x = iris, tableName = "iris")
#' write_datatable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#'
#' ## wb_remove_tables() deletes table object and all data
#' wb_get_tables(wb, sheet = 1)
#' wb$remove_tables(sheet = 1, table = "iris")
#' write_datatable(wb, sheet = 1, x = iris, tableName = "iris", startCol = 1)
#'
#' wb_get_tables(wb, sheet = 1)
#' wb$remove_tables(sheet = 1, table = "iris")
#' write_datatable(wb, sheet = 1, x = iris, tableName = "iris", startCol = 1)
#' \dontrun{
#' wb_save(wb, path = "wb_remove_tablesExample.xlsx", overwrite = TRUE)
#' }
#'
#' @export
wb_remove_tables <- function(wb, sheet, table) {
  assert_workbook(wb)
  wb$clone()$remove_tables(sheet = sheet, table = table)
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
#' wb$add_worksheet("AirPass")
#' write_data(wb, "AirPass", t2, rowNames = TRUE)
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
  wb$clone()$group_cols(
    sheet     = sheet,
    cols      = cols,
    collapsed = collapsed,
    levels    = levels
  )
}

#' @export
#' @rdname workbook_grouping
wb_ungroup_cols <- function(wb, sheet, cols) {
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
  wb$clone()$group_rows(
    sheet     = sheet,
    rows      = rows,
    collapsed = collapsed,
    levels    = levels
  )
}

#' @export
#' @rdname workbook_grouping
wb_ungroup_rows <- function(wb, sheet, rows) {
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
  wb$clone()$add_creators(creators)
}

#' @rdname wb_creators
#' @export
wb_set_creators <- function(wb, creators) {
  assert_workbook(wb)
  wb$clone()$set_creators(creators)
}

#' @rdname wb_creators
#' @export
wb_remove_creators <- function(wb, creators) {
  assert_workbook(wb)
  wb$clone()$remove_creators(creators)
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
#' Just a wrapper of wb$set_last_modified_by()
#'
#' @param wb A workbook object
#' @param LastModifiedBy A string object with the name of the LastModifiedBy-User
#'
#' @export
#' @family workbook wrappers
#'
#' @examples
#' wb <- wb_workbook()
#' wb_set_last_modified_by(wb, "test")
wb_set_last_modified_by <- function(wb, LastModifiedBy) {
  assert_workbook(wb)
  wb$clone()$set_last_modified_by(LastModifiedBy)
}

#' Insert an image into a worksheet
#'
#' Insert an image into a worksheet
#'
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param file An image file. Valid file types are:` "jpeg"`, `"png"`, `"bmp"`
#' @param width Width of figure.
#' @param height Height of figure.
#' @param startRow Row coordinate of upper left corner of the image
#' @param startCol Column coordinate of upper left corner of the image
#' @param rowOffset offset within cell (row)
#' @param colOffset offset within cell (column)
#' @param units Units of width and height. Can be `"in"`, `"cm"` or `"px"`
#' @param dpi Image resolution used for conversion between units.
#' @seealso [wb_add_plot()]
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- wb_workbook("Ayanami")
#'
#' ## Add some worksheets
#' wb$add_worksheet("Sheet 1")
#' wb$add_worksheet("Sheet 2")
#' wb$add_worksheet("Sheet 3")
#'
#' ## Insert images
#' img <- system.file("extdata", "einstein.jpg", package = "openxlsx2")
#' wb$add_image("Sheet 1", img, startRow = 5, startCol = 3, width = 6, height = 5)
#' wb$add_image(2, img, startRow = 2, startCol = 2)
#' wb$add_image(3, img, width = 15, height = 12, startRow = 3, startCol = "G", units = "cm")
#'
#' ## Save workbook
#' \dontrun{
#' wb_save(wb, "wb_add_imageExample.xlsx", overwrite = TRUE)
#' }
wb_add_image <- function(
  wb,
  sheet,
  file,
  width     = 6,
  height    = 3,
  startRow  = 1,
  startCol  = 1,
  rowOffset = 0,
  colOffset = 0,
  units     = "in",
  dpi       = 300
) {
  assert_workbook(wb)
  wb$clone()$add_image(
    sheet     = sheet,
    file      = file,
    startRow  = startRow,
    startCol  = startCol,
    width     = width,
    height    = height,
    rowOffset = rowOffset,
    colOffset = colOffset,
    units     = units,
    dpi       = dpi
  )
}
