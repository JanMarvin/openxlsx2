

# R6 class ----------------------------------------------------------------

#' R6 class for a Workbook
#'
#' A Workbook
#'
#' @export
wbWorkbook <- R6::R6Class(
  "wbWorkbook",

  # TODO which can be private?

  ## public ----

  public = list(
    #' @field sheet_names sheet_names
    sheet_names = character(),

    #' @field calcChain calcChain
    calcChain = character(),

    #' @field apps apps
    apps = character(),

    #' @field charts charts
    charts = list(),

    #' @field isChartSheet isChartSheet
    isChartSheet = logical(),

    #' @field customXml customXml
    customXml = NULL,

    #' @field connections connections
    connections = NULL,

    #' @field ctrlProps ctrlProps
    ctrlProps = NULL,

    #' @field Content_Types Content_Types
    Content_Types = genBaseContent_Type(),

    #' @field app app
    app = character(),

    #' @field core core
    core = character(),

    #' @field custom custom
    custom = character(),

    #' @field drawings drawings
    drawings = NULL,

    #' @field drawings_rels drawings_rels
    drawings_rels = NULL,

    # #' @field drawings_vml drawings_vml
    # drawings_vml = NULL,

    #' @field embeddings embeddings
    embeddings = NULL,

    #' @field externalLinks externalLinks
    externalLinks = NULL,

    #' @field externalLinksRels externalLinksRels
    externalLinksRels = NULL,

    #' @field headFoot headFoot
    headFoot = NULL,

    #' @field media media
    media = NULL,

    #' @field metadata metadata
    metadata = NULL,

    #' @field persons persons
    persons = NULL,

    #' @field pivotTables pivotTables
    pivotTables = NULL,

    #' @field pivotTables.xml.rels pivotTables.xml.rels
    pivotTables.xml.rels = NULL,

    #' @field pivotDefinitions pivotDefinitions
    pivotDefinitions = NULL,

    #' @field pivotRecords pivotRecords
    pivotRecords = NULL,

    #' @field pivotDefinitionsRels pivotDefinitionsRels
    pivotDefinitionsRels = NULL,

    #' @field queryTables queryTables
    queryTables = NULL,

    #' @field slicers slicers
    slicers = NULL,

    #' @field slicerCaches slicerCaches
    slicerCaches = NULL,

    #' @field sharedStrings sharedStrings
    sharedStrings = structure(list(), uniqueCount = 0L),

    #' @field styles_mgr styles_mgr
    styles_mgr = NULL,

    #' @field styles_xml styles_xml
    styles_xml = NULL,

    #' @field tables tables
    tables = NULL,

    #' @field tables.xml.rels tables.xml.rels
    tables.xml.rels = NULL,

    #' @field theme theme
    theme = NULL,

    #' @field vbaProject vbaProject
    vbaProject = NULL,

    #' @field vml vml
    vml = list(),

    #' @field vml_rels vml_rels
    vml_rels = list(),

    #' @field comments comments
    comments = list(),

    #' @field threadComments threadComments
    threadComments = NULL,

    #' @field workbook workbook
    workbook = genBaseWorkbook(),

    #' @field workbook.xml.rels workbook.xml.rels
    workbook.xml.rels = genBaseWorkbook.xml.rels(),

    #' @field worksheets worksheets
    worksheets = list(),

    #' @field worksheets_rels worksheets_rels
    worksheets_rels = list(),

    #' @field sheetOrder The sheet order.  Controls ordering for worksheets and
    #'   worksheet names.
    sheetOrder = integer(),

    #' @field path path
    path = character(),     # allows path to be set during initiation or later

    #' @field creator A character vector of creators
    creator = character(),

    #' @field title title
    title = NULL,

    #' @field subject subject
    subject = NULL,

    #' @field category category
    category = NULL,

    #' @field datetimeCreated The datetime (as `POSIXt`) the workbook is
    #'   created.  Defaults to the current `Sys.time()` when the workbook object
    #'   is created, not when the Excel files are saved.
    datetimeCreated = Sys.time(),

    #' @description
    #' Creates a new `wbWorkbook` object
    #' @param creator character vector of creators.  Duplicated are ignored.
    #' @param title title
    #' @param subject subject
    #' @param category category
    #' @param datetimeCreated The datetime (as `POSIXt`) the workbook is
    #'   created.  Defaults to the current `Sys.time()` when the workbook object
    #'   is created, not when the Excel files are saved.
    #' @return a `wbWorkbook` object
    initialize = function(
    creator         = NULL,
    title           = NULL,
    subject         = NULL,
    category        = NULL,
    datetimeCreated = Sys.time()
    ) {
      workbook_initialize(
        self            = self,
        private         = private,
        creator         = creator,
        title           = title,
        subject         = subject,
        category        = category,
        datetimeCreated = datetimeCreated
      )
    },

    #' @description
    #' Append a field. This method is intended for internal use
    #' @param field A valid field name
    #' @param value A value for the field
    append = function(field, value) {
      workbook_append(self, private, field, value)
    },

    #' @description
    #' Append to `self$workbook$sheets` This method is intended for internal use
    #' @param value A value for `self$workbook$sheets`
    append_sheets = function(value) {
      workbook_append_sheets(self, private, value)
    },

    #' @description validate sheet
    #' @param sheet A character sheet name or integer location
    #' @returns The integer position of the sheet
    validate_sheet = function(sheet) {
      workbook_validate_sheet(self, private, sheet)
    },

    #' @description
    #' Add worksheet to the `wbWorkbook` object
    #' @param sheet sheet
    #' @param gridLines gridLines
    #' @param rowColHeaders rowColHeaders
    #' @param tabColour tabColour
    #' @param zoom zoom
    #' @param header header
    #' @param footer footer
    #' @param oddHeader oddHeader
    #' @param oddFooter oddFooter
    #' @param evenHeader evenHeader
    #' @param evenFooter evenFooter
    #' @param firstHeader firstHeader
    #' @param firstFooter firstFooter
    #' @param visible visible
    #' @param hasDrawing hasDrawing
    #' @param paperSize paperSize
    #' @param orientation orientation
    #' @param hdpi hdpi
    #' @param vdpi vdpi
    #' @return The `wbWorkbook` object, invisibly
    add_worksheet = function(
      sheet       = next_sheet(),
      gridLines   = TRUE,
      rowColHeaders = TRUE,
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
      workbook_add_worksheet(
        self          = self,
        private       = private,
        sheet         = sheet,
        gridLines     = gridLines,
        rowColHeaders = rowColHeaders,
        tabColour     = tabColour,
        zoom          = zoom,
        header        = header,
        footer        = footer,
        oddHeader     = oddHeader,
        oddFooter     = oddFooter,
        evenHeader    = evenHeader,
        evenFooter    = evenFooter,
        firstHeader   = firstHeader,
        firstFooter   = firstFooter,
        visible       = visible,
        hasDrawing    = hasDrawing,
        paperSize     = paperSize,
        orientation   = orientation,
        hdpi          = hdpi,
        vdpi          = vdpi
      )
    },

    # TODO should this be as simple as: wb$wb_add_worksheet(wb$worksheets[[1]]$clone()) ?

    #' @description
    #' Clone a workbooksheet
    #' @param old name of worksheet to clone
    #' @param new name of new worksheet to add
    clone_worksheet = function(old = current_sheet(), new = next_sheet()) {
      workbook_clone_worksheet(self, private, old = old, new = new)
    },

    #' @description
    #' Add a chart sheet to the workbook
    #' @param sheet sheet
    #' @param tabColour tabColour
    #' @param zoom zoom
    #' @return The `wbWorkbook` object, invisibly
    addChartSheet = function(sheet = current_sheet(), tabColour = NULL, zoom = 100) {
      workbook_add_chartsheet(
        self      = self,
        private   = private,
        sheet     = sheet,
        tabColour = tabColour,
        zoom      = zoom
      )
    },

    ### add data ----

    #' @description add data
    #' @param sheet sheet
    #' @param x x
    #' @param startCol startCol
    #' @param startRow startRow
    #' @param dims dims
    #' @param array array
    #' @param xy xy
    #' @param colNames colNames
    #' @param rowNames rowNames
    #' @param withFilter withFilter
    #' @param name name
    #' @param sep sep
    #' @param applyCellStyle applyCellStyle
    #' @param removeCellStyle if writing into existing cells, should the cell style be removed?
    #' @param na.strings na.strings
    #' @param return The `wbWorkbook` object
    add_data = function(
      sheet           = current_sheet(),
      x,
      startCol        = 1,
      startRow        = 1,
      dims            = rowcol_to_dims(startRow, startCol),
      array           = FALSE,
      xy              = NULL,
      colNames        = TRUE,
      rowNames        = FALSE,
      withFilter      = FALSE,
      name            = NULL,
      sep             = ", ",
      applyCellStyle  = TRUE,
      removeCellStyle = FALSE,
      na.strings      = getOption("openxlsx2.na.strings", "#N/A")
    ) {
      # TODO shouldn't this have a default?
      workbook_add_data(
        self            = self,
        private         = private,
        sheet           = sheet,
        x               = x,
        startCol        = startCol,
        startRow        = startRow,
        dims            = dims,
        array           = array,
        xy              = xy,
        colNames        = colNames,
        rowNames        = rowNames,
        withFilter      = withFilter,
        name            = name,
        sep             = sep,
        applyCellStyle  = applyCellStyle,
        removeCellStyle = removeCellStyle,
        na.strings      = na.strings
      )
    },

    #' @description add a data table
    #' @param sheet sheet
    #' @param x x
    #' @param startCol startCol
    #' @param startRow startRow
    #' @param dims dims
    #' @param xy xy
    #' @param colNames colNames
    #' @param rowNames rowNames
    #' @param tableStyle tableStyle
    #' @param tableName tableName
    #' @param withFilter withFilter
    #' @param sep sep
    #' @param firstColumn firstColumn
    #' @param lastColumn lastColumn
    #' @param bandedRows bandedRows
    #' @param bandedCols bandedCols
    #' @param applyCellStyle applyCellStyle
    #' @param removeCellStyle if writing into existing cells, should the cell style be removed?
    #' @param na.strings na.strings
    #' @returns The `wbWorkbook` object
    add_data_table = function(
      sheet           = current_sheet(),
      x,
      startCol        = 1,
      startRow        = 1,
      dims            = rowcol_to_dims(startRow, startCol),
      xy              = NULL,
      colNames        = TRUE,
      rowNames        = FALSE,
      tableStyle      = "TableStyleLight9",
      tableName       = NULL,
      withFilter      = TRUE,
      sep             = ", ",
      firstColumn     = FALSE,
      lastColumn      = FALSE,
      bandedRows      = TRUE,
      bandedCols      = FALSE,
      applyCellStyle  = TRUE,
      removeCellStyle = FALSE,
      na.strings      = getOption("openxlsx2.na.strings", "#N/A")
    ) {
      workbook_add_data_table(
        self            = self,
        private         = private,
        sheet           = sheet,
        x               = x,
        startCol        = startCol,
        startRow        = startRow,
        dims            = dims,
        xy              = xy,
        colNames        = colNames,
        rowNames        = rowNames,
        tableStyle      = tableStyle,
        tableName       = tableName,
        withFilter      = withFilter,
        sep             = sep,
        firstColumn     = firstColumn,
        lastColumn      = lastColumn,
        bandedRows      = bandedRows,
        bandedCols      = bandedCols,
        applyCellStyle  = applyCellStyle,
        removeCellStyle = removeCellStyle,
        na.strings      = na.strings
      )
    },
    #' @description add formula
    #' @param sheet sheet
    #' @param x x
    #' @param startCol startCol
    #' @param startRow startRow
    #' @param dims dims
    #' @param array array
    #' @param xy xy
    #' @param applyCellStyle applyCellStyle
    #' @param removeCellStyle if writing into existing cells, should the cell style be removed?
    #' @returns The `wbWorkbook` object
    add_formula = function(
      sheet    = current_sheet(),
      x,
      startCol = 1,
      startRow = 1,
      dims     = rowcol_to_dims(startRow, startCol),
      array    = FALSE,
      xy       = NULL,
      applyCellStyle = TRUE,
      removeCellStyle = FALSE
    ) {
      workbook_add_formula(
        self            = self,
        private         = private,
        sheet           = sheet,
        x               = x,
        startCol        = startCol,
        startRow        = startRow,
        dims            = dims,
        array           = array,
        xy              = xy,
        applyCellStyle  = applyCellStyle,
        removeCellStyle = removeCellStyle
      )
    },

    #' @description add style
    #' @param style style
    #' @param style_name style_name
    #' @returns The `wbWorkbook` object
    add_style = function(style = NULL, style_name = substitute(style)) {
      force(style_name)
      workbook_add_style(self, private, style = style, style_name = style_name)
    },

    # TODO wb_save can be shortened a lot by some formatting and by using a
    # function that creates all the temporary directories and subdirectries as a
    # named list

    #' @description
    #' Save the workbook
    #' @param path The path to save the workbook to
    #' @param overwrite If `FALSE`, will not overwrite when `path` exists
    #' @return The `wbWorkbook` object invisibly
    save = function(path = self$path, overwrite = TRUE) {
      workbook_save(self, private, path = path, overwrite = overwrite)
    },

    #' @description open wbWorkbook in Excel.
    #' @details minor helper wrapping xl_open which does the entire same thing
    #' @param interactive If `FALSE` will throw a warning and not open the path.
    #'   This can be manually set to `TRUE`, otherwise when `NA` (default) uses
    #'   the value returned from [base::interactive()]
    #' @return The `wbWorkbook`, invisibly
    open = function(interactive = NA) {
      workbook_open(self, private, interactive)
    },

    #' @description
    #' Build table
    #' @param sheet sheet
    #' @param colNames colNames
    #' @param ref ref
    #' @param showColNames showColNames
    #' @param tableStyle tableStyle
    #' @param tableName tableName
    #' @param withFilter withFilter
    #' @param totalsRowCount totalsRowCount
    #' @param showFirstColumn showFirstColumn
    #' @param showLastColumn showLastColumn
    #' @param showRowStripes showRowStripes
    #' @param showColumnStripes showColumnStripes
    #' @return The `wbWorksheet` object, invisibly
    buildTable = function(
      sheet = current_sheet(),
      colNames,
      ref,
      showColNames,
      tableStyle,
      tableName,
      withFilter, # TODO set default for withFilter?
      totalsRowCount = 0,
      showFirstColumn = 0,
      showLastColumn = 0,
      showRowStripes = 1,
      showColumnStripes = 0
    ) {
      workbook_build_table(
        self              = self,
        private           = private,
        sheet             = sheet,
        colNames          = colNames,
        ref               = ref,
        showColNames      = showColNames,
        tableStyle        = tableStyle,
        tableName         = tableName,
        withFilter        = withFilter,
        totalsRowCount    = totalsRowCount,
        showFirstColumn   = showFirstColumn,
        showLastColumn    = showLastColumn,
        showRowStripes    = showRowStripes,
        showColumnStripes = showColumnStripes
      )
    },

    ### base font ----

    #' @description
    #' Get the base font
    #' @return A list of of the font
    get_base_font = function() {
      workbook_get_base_font(self, private)
    },

    #' @description
    #' Get the base font
    #' @param fontSize fontSize
    #' @param fontColour fontColour
    #' @param fontName fontName
    #' @return The `wbWorkbook` object
    set_base_font = function(
      fontSize = 11,
      fontColour = wb_colour(theme = "1"),
      fontName = "Calibri"
    ) {
      workbook_set_base_font(
        self       = self,
        private    = private,
        fontSize   = fontSize,
        fontColour = fontColour,
        fontName   = fontName
      )
    },

    ### book views ----

    #' @description
    #' Set the book views
    #' @param activeTab activeTab
    #' @param autoFilterDateGrouping autoFilterDateGrouping
    #' @param firstSheet firstSheet
    #' @param minimized minimized
    #' @param showHorizontalScroll showHorizontalScroll
    #' @param showSheetTabs showSheetTabs
    #' @param showVerticalScroll showVerticalScroll
    #' @param tabRatio tabRatio
    #' @param visibility visibility
    #' @param windowHeight windowHeight
    #' @param windowWidth windowWidth
    #' @param xWindow xWindow
    #' @param yWindow yWindow
    #' @return The `wbWorkbook` object
    set_bookview = function(
      activeTab              = NULL,
      autoFilterDateGrouping = NULL,
      firstSheet             = NULL,
      minimized              = NULL,
      showHorizontalScroll   = NULL,
      showSheetTabs          = NULL,
      showVerticalScroll     = NULL,
      tabRatio               = NULL,
      visibility             = NULL,
      windowHeight           = NULL,
      windowWidth            = NULL,
      xWindow                = NULL,
      yWindow                = NULL
    ) {
      workbook_set_bookview(
        self                   = self,
        private                = private,
        activeTab              = activeTab,
        autoFilterDateGrouping = autoFilterDateGrouping,
        firstSheet             = firstSheet,
        minimized              = minimized,
        showHorizontalScroll   = showHorizontalScroll,
        showSheetTabs          = showSheetTabs,
        showVerticalScroll     = showVerticalScroll,
        tabRatio               = tabRatio,
        visibility             = visibility,
        windowHeight           = windowHeight,
        windowWidth            = windowWidth,
        xWindow                = xWindow,
        yWindow                = yWindow
      )
    },

    ### sheet names ----

    #' @description Get sheet names
    #' @returns A `named` `character` vector of sheet names in their order.  The
    #'   names represent the original value of the worksheet prior to any
    #'   character substitutions.
    get_sheet_names = function() {
      workbook_get_sheet_names(self, private)
    },

    #' @description
    #' Sets a sheet name
    #' @param old Old sheet name
    #' @param new New sheet name
    #' @return The `wbWorkbook` object, invisibly
    set_sheet_names = function(old = NULL, new) {
      workbook_set_sheet_names(self, private, old, new)
    },

    #' @description
    #' Deprecated.  Use `set_sheet_names()` instead
    #' @param sheet Old sheet name
    #' @param name New sheet name
    #' @return The `wbWorkbook` object, invisibly
    setSheetName = function(sheet = current_sheet(), name) {
      .Deprecated("wbWorkbook$set_sheet_names()")
      self$set_sheet_names(old = sheet, new = name)
    },

    ### row heights ----

    #' @description
    #' Sets a row height for a sheet
    #' @param sheet sheet
    #' @param rows rows
    #' @param heights heights
    #' @return The `wbWorkbook` object, invisibly
    set_row_heights = function(sheet = current_sheet(), rows, heights) {
      workbook_set_row_heights(
        self    = self,
        private = private,
        sheet   = sheet,
        rows    = rows,
        heights = heights
      )
    },

    #' @description
    #' Sets a row height for a sheet
    #' @param sheet sheet
    #' @param rows rows
    #' @return The `wbWorkbook` object, invisibly
    remove_row_heights = function(sheet = current_sheet(), rows) {
      workbook_remove_row_heights(
        self    = self,
        private = private,
        sheet   = sheet,
        rows    = rows
      )
    },

    ### columns ----

    #' description
    #' creates column object for worksheet
    #' @param sheet sheet
    #' @param n n
    #' @param beg beg
    #' @param end end
    createCols = function(sheet = current_sheet(), n, beg, end) {
      # TODO replace with $add_cols()
      workbook_add_cols(
        self    = self,
        private = private,
        sheet   = sheet,
        n       = n,
        beg     = beg,
        end     = end
      )
    },

    #' @description
    #' Group cols
    #' @param sheet sheet
    #' @param cols cols
    #' @param collapsed collapsed
    #' @param levels levels
    #' @return The `wbWorkbook` object, invisibly
    group_cols = function(
      sheet = current_sheet(),
      cols,
      collapsed = FALSE,
      levels = NULL
    ) {
      workbook_group_cols(
        self      = self,
        private   = private,
        sheet     = sheet,
        cols      = cols,
        collapsed = collapsed,
        levels    = levels
      )
    },

    #' @description ungroup cols
    #' @param sheet sheet
    #' @param cols = cols
    #' @returns The `wbWorkbook` object
    ungroup_cols = function(sheet = current_sheet(), cols) {
      workbook_ungroup_cols(self, private, sheet, cols)
    },

    #' @description Remove row heights from a worksheet
    #' @param sheet A name or index of a worksheet
    #' @param cols Indices of columns to remove custom width (if any) from.
    #' @return The `wbWorkbook` object, invisibly
    remove_col_widths = function(sheet = current_sheet(), cols) {
      workbook_remove_col_widths(
        self    = self,
        private = private,
        sheet   = sheet,
        cols    = cols
      )
    },

    # TODO wb_group_rows() and group_cols() are very similiar.  Can problem turn
    #' @description
    #' Group cols
    #' @param sheet sheet
    #' @param cols cols
    #' @param widths Width of columns
    #' @param hidden A logical vector to determine which cols are hidden; values
    #'   are repeated across length of `cols`
    #' @return The `wbWorkbook` object, invisibly
    set_col_widths = function(
      sheet = current_sheet(),
      cols,
      widths = 8.43,
      hidden = FALSE
    ) {
      workbook_set_col_widths(
        self    = self,
        private = private,
        sheet   = sheet,
        cols    = cols,
        widths  = widths,
        hidden  = hidden
      )
    },

    ### rows ----

    # TODO groupRows() and groupCols() are very similiar.  Can problem turn
    # these into some wrappers for another method

    #' @description
    #' Group rows
    #' @param sheet sheet
    #' @param rows rows
    #' @param collapsed collapsed
    #' @param levels levels
    #' @return The `wbWorkbook` object, invisibly
    group_rows = function(
      sheet = current_sheet(),
      rows,
      collapsed = FALSE,
      levels = NULL
    ) {
      workbook_group_rows(
        self      = self,
        private   = private,
        sheet     = sheet,
        rows      = rows,
        collapsed = collapsed,
        levels    = levels
      )
    },

    #' @description ungroup rows
    #' @param sheet sheet
    #' @param rows rows
    #' @return The `wbWorkbook` object
    ungroup_rows = function(sheet = current_sheet(), rows) {
      workbook_ungroup_rows(self, private, sheet = sheet, rows = rows)
    },

    #' @description
    #' Remove a worksheet
    #' @param sheet The worksheet to delete
    #' @return The `wbWorkbook` object, invisibly
    remove_worksheet = function(sheet = current_sheet()) {
      workbook_remove_worksheet(self, private, sheet = sheet)
    },

    ### cells ----

    #' @description Adds data validation
    #' @param sheet sheet
    #' @param cols cols
    #' @param rows rows
    #' @param type type
    #' @param operator operator
    #' @param value value
    #' @param allowBlank allowBlank
    #' @param showInputMsg showInputMsg
    #' @param showErrorMsg showErrorMsg
    #' @param errorStyle The icon shown and the options how to deal with such inputs. Default "stop" (cancel), else "information" (prompt popup) or "warning" (prompt accept or change input)
    #' @param errorTitle The error title
    #' @param error The error text
    #' @param promptTitle The prompt title
    #' @param prompt The prompt text
    #' @returns The `wbWorkbook` object
    add_data_validation = function(
      sheet        = current_sheet(),
      cols,
      rows,
      type,
      operator,
      value,
      allowBlank   = TRUE,
      showInputMsg = TRUE,
      showErrorMsg = TRUE,
      errorStyle   = NULL,
      errorTitle   = NULL,
      error        = NULL,
      promptTitle  = NULL,
      prompt       = NULL
    ) {
      workbook_add_data_validations(
        self         = self,
        private      = private,
        sheet        = sheet,
        cols         = cols,
        rows         = rows,
        type         = type,
        operator     = operator,
        value        = value,
        allowBlank   = allowBlank,
        showInputMsg = showInputMsg,
        showErrorMsg = showErrorMsg,
        errorStyle   = errorStyle,
        errorTitle   = errorTitle,
        error        = error,
        promptTitle  = promptTitle,
        prompt       = prompt
      )
    },

    #' @description
    #' Set cell merging for a sheet
    #' @param sheet sheet
    #' @param rows,cols Row and column specifications.
    #' @return The `wbWorkbook` object, invisibly
    merge_cells = function(sheet = current_sheet(), rows = NULL, cols = NULL) {
      workbook_merge_cells(
        self    = self,
        private = private,
        sheet   = sheet,
        rows    = rows,
        cols    = cols
      )
    },

    #' @description
    #' Removes cell merging for a sheet
    #' @param sheet sheet
    #' @param rows,cols Row and column specifications.
    #' @return The `wbWorkbook` object, invisibly
    unmerge_cells = function(
      sheet = current_sheet(),
      rows = NULL,
      cols = NULL
    ) {
      workbook_unmerge_cells(
        self    = self,
        private = private,
        sheet   = sheet,
        rows    = rows,
        cols    = cols
      )
    },

    #' @description
    #' Set freeze panes for a sheet
    #' @param sheet sheet
    #' @param firstActiveRow firstActiveRow
    #' @param firstActiveCol firstActiveCol
    #' @param firstRow firstRow
    #' @param firstCol firstCol
    #' @return The `wbWorkbook` object, invisibly
    freeze_pane = function(
      sheet          = current_sheet(),
      firstActiveRow = NULL,
      firstActiveCol = NULL,
      firstRow       = FALSE,
      firstCol       = FALSE
    ) {
      workbook_free_panes(
        self           = self,
        private        = private,
        sheet          = sheet,
        firstActiveRow = firstActiveRow,
        firstActiveCol = firstActiveCol,
        firstRow       = firstRow,
        firstCol       = firstCol
      )
    },

    ### comment ----

    #' @description Add comment
    #' @param sheet sheet
    #' @param col column to apply the comment
    #' @param row row to apply the comment
    #' @param dims row and column as spreadsheet dimension, e.g. "A1"
    #' @param comment a comment to apply to the worksheet
    #' @returns The `wbWorkbook` object
    add_comment = function(
      sheet = current_sheet(),
      col,
      row,
      dims  = rowcol_to_dims(row, col),
      comment
    ) {
      workbook_add_comment(
        self    = self,
        private = private,
        sheet   = sheet,
        col     = col,
        row     = row,
        dims    = dims,
        comment = comment
      )
    },

    #' @description Remove comment
    #' @param sheet sheet
    #' @param col column to apply the comment
    #' @param row row to apply the comment
    #' @param dims row and column as spreadsheet dimension, e.g. "A1"
    #' @param gridExpand Remove all comments inside the grid. Similar to dims "A1:B2"
    #' @returns The `wbWorkbook` object
    remove_comment = function(
      sheet      = current_sheet(),
      col,
      row,
      dims       = rowcol_to_dims(row, col),
      gridExpand = TRUE
    ) {
      workbook_remove_comment(
        self       = self,
        private    = private,
        sheet      = sheet,
        col        = col,
        row        = row,
        dims       = dims,
        gridExpand = gridExpand
      )
    },

    ### conditional formatting ----

    # TODO remove_conditional_formatting?

    #' @description Add conditional formatting
    #' @param sheet sheet
    #' @param cols cols
    #' @param rows rows
    #' @param rule rule
    #' @param style style
    #' @param type type
    #' @param params Additional parameters
    #' @returns The `wbWorkbook` object
    add_conditional_formatting = function(
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
      workbook_add_conditional_formatting(
        self    = self,
        private = private,
        sheet   = sheet,
        cols    = cols,
        rows    = rows,
        rule    = rule,
        style   = style,
        type    = type,
        params  = params
      )
    },

    ### plots and images ----

    #' @description
    #' Insert an image into a sheet
    #' @param sheet sheet
    #' @param file file
    #' @param startRow startRow
    #' @param startCol startCol
    #' @param width width
    #' @param height height
    #' @param rowOffset rowOffset
    #' @param colOffset colOffset
    #' @param units units
    #' @param dpi dpi
    #' @return The `wbWorkbook` object, invisibly
    add_image = function(
      sheet     = current_sheet(),
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
      workbook_add_image(
        self      = self,
        private   = private,
        sheet     = sheet,
        file      = file,
        width     = width,
        height    = height,
        startRow  = startRow,
        startCol  = startCol,
        rowOffset = rowOffset,
        colOffset = colOffset,
        units     = units,
        dpi       = dpi
      )
    },

    #' @description Add plot. A wrapper for add_image()
    #' @param sheet sheet
    #' @param width width
    #' @param height height
    #' @param xy xy
    #' @param startRow startRow
    #' @param startCol startCol
    #' @param rowOffset rowOffset
    #' @param colOffset colOffset
    #' @param fileType fileType
    #' @param units units
    #' @param dpi dpi
    #' @returns The `wbWorkbook` object
    add_plot = function(
      sheet = current_sheet(),
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
      workbook_add_plot(
        self      = self,
        private   = private,
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
    },

    #' @description Add xml drawing
    #' @param sheet sheet
    #' @param dims dims
    #' @param xml xml
    #' @returns The `wbWorkbook` object
    add_drawing = function(
      sheet = current_sheet(),
      xml,
      dims  = "A1:H8"
    ) {
      workbook_add_drawing(
        self    = self,
        private = private,
        sheet   = sheet,
        xml     = xml,
        dims    = dims
      )
    },

    #' @description Add xml drawing
    #' @description Add xml chart
    #' @param sheet sheet
    #' @param dims dims
    #' @param xml xml
    #' @returns The `wbWorkbook` object
    add_chart_xml = function(
      sheet = current_sheet(),
      xml,
      dims  = "A1:H8"
    ) {
      workbook_add_chart_xml(
        self    = self,
        private = private,
        sheet   = sheet,
        xml     = xml,
        dims    = dims
      )
    },

    #' @description Add mschart chart to the workbook
    #' @param sheet the sheet on which the graph will appear
    #' @param dims the dimensions where the sheet will appear
    #' @param graph mschart graph
    #' @returns The `wbWorkbook` object
    add_mschart = function(
      sheet = current_sheet(),
      dims = "B2:H8",
      graph
    ) {
      workbook_add_mschart(
        self    = self,
        private = private,
        sheet   = sheet,
        dims    = dims,
        graph   = graph
      )
    },

    ### methods ----

    #' @description
    #' Prints the `wbWorkbook` object
    #' @return The `wbWorkbook` object, invisibly; called for its side-effects
    print = function() {
      workbook_print(self, private)
    },

    ### protect ---

    #' @description
    #' Protect a workbook
    #' @param protect protect
    #' @param lockStructure lockStructure
    #' @param lockWindows lockWindows
    #' @param password password
    #' @param type type
    #' @param fileSharing fileSharing
    #' @param username username
    #' @param readOnlyRecommended readOnlyRecommended
    #' @return The `wbWorkbook` object, invisibly
    protect = function(
      protect             = TRUE,
      password            = NULL,
      lockStructure       = FALSE,
      lockWindows         = FALSE,
      type                = c("1", "2", "4", "8"),
      fileSharing         = FALSE,
      username            = unname(Sys.info()["user"]),
      readOnlyRecommended = FALSE
    ) {
      workbook_protect(
        self                = self,
        private             = private,
        protect             = protect,
        password            = password,
        lockStructure       = lockStructure,
        lockWindows         = lockWindows,
        type                = type,
        fileSharing         = fileSharing,
        username            = username,
        readOnlyRecommended = readOnlyRecommended
      )
    },

    #' @description protect worksheet
    #' @param sheet sheet
    #' @param protect protect
    #' @param password password
    #' @param properties A character vector of properties to lock.  Can be one
    #'   or more of the following: `"selectLockedCells"`,
    #'   `"selectUnlockedCells"`, `"formatCells"`, `"formatColumns"`,
    #'   `"formatRows"`, `"insertColumns"`, `"insertRows"`,
    #'   `"insertHyperlinks"`, `"deleteColumns"`, `"deleteRows"`, `"sort"`,
    #'   `"autoFilter"`, `"pivotTables"`, `"objects"`, `"scenarios"`
    #' @returns The `wbWorkbook` object
    protect_worksheet = function(
      sheet = current_sheet(),
      protect    = TRUE,
      password   = NULL,
      properties = NULL
    ) {
      workbook_protect_worksheet(
        self       = self,
        private    = private,
        sheet      = sheet,
        protect    = protect,
        password   = password,
        properties = properties
      )
    },

    ### creators --------------------------------------------------------------

    #' @description Set creator(s)
    #' @param creators A character vector of creators to set.  Duplicates are
    #'   ignored.
    set_creators = function(creators) {
      workbook_modify_creators(self, private, "set", creators)
    },

    #' @description Add creator(s)
    #' @param creators A character vector of creators to add.  Duplicates are
    #'   ignored.
    add_creators = function(creators) {
      workbook_modify_creators(self, private, "add", creators)
    },

    #' @description Remove creator(s)
    #' @param creators A character vector of creators to remove.  All duplicated
    #'   are removed.
    remove_creators = function(creators) {
      workbook_modify_creators(self, private, "remove", creators)
    },

    ### last modified by ----

    #' @description
    #' Change the last modified by
    #' @param LastModifiedBy A new value
    #' @return The `wbWorkbook` object, invisibly
    set_last_modified_by = function(LastModifiedBy = NULL) {
      workbook_set_last_modified_by(self, private, LastModifiedBy)
    },

    ### page setup ----

    #' @description page_setup()
    #' @param sheet sheet
    #' @param orientation orientation
    #' @param scale scale
    #' @param left left
    #' @param right right
    #' @param top top
    #' @param bottom bottom
    #' @param header header
    #' @param footer footer
    #' @param fitToWidth fitToWidth
    #' @param fitToHeight fitToHeight
    #' @param paperSize paperSize
    #' @param printTitleRows printTitleRows
    #' @param printTitleCols printTitleCols
    #' @param summaryRow summaryRow
    #' @param summaryCol summaryCol
    #' @return The `wbWorkbook` object, invisibly
    page_setup = function(
      sheet = current_sheet(),
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
      workbook_page_setup(
        self           = self,
        private        = private,
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
    },

    ### header footer ----

    #' @description Sets headers and footers
    #' @param sheet sheet
    #' @param header header
    #' @param footer footer
    #' @param evenHeader evenHeader
    #' @param evenFooter evenFooter
    #' @param firstHeader firstHeader
    #' @param firstFooter firstFooter
    #' @return The `wbWorkbook` object, invisibly
    set_header_footer = function(
      sheet = current_sheet(),
      header      = NULL,
      footer      = NULL,
      evenHeader  = NULL,
      evenFooter  = NULL,
      firstHeader = NULL,
      firstFooter = NULL
    ) {
      workbook_set_header_footer(
        self        = self,
        private     = private,
        sheet       = sheet,
        header      = header,
        footer      = footer,
        evenHeader  = evenHeader,
        evenFooter  = evenFooter,
        firstHeader = firstHeader,
        firstFooter = firstFooter
      )
    },

    ### tables ----

    #' @description get tables
    #' @param sheet sheet
    #' @returns The sheet tables.  `character()` if empty
    get_tables = function(sheet = current_sheet()) {
      workbook_get_tables(
        self    = self,
        private = private,
        sheet   = sheet
      )
    },

    #' @description remove tables
    #' @param sheet sheet
    #' @param table table
    #' @returns The `wbWorkbook` object
    remove_tables = function(sheet = current_sheet(), table) {
      workbook_remove_tables(
        self    = self,
        private = private,
        sheet   = sheet,
        table   = table
      )
    },

    ### filters ----

    #' @description add filters
    #' @param sheet sheet
    #' @param rows rows
    #' @param cols cols
    #' @returns The `wbWorkbook` object
    add_filter = function(sheet = current_sheet(), rows, cols) {
      workbook_add_filter(
        self    = self,
        private = private,
        sheet   = sheet,
        rows    = rows,
        cols    = cols
      )
    },

    #' @description remove filters
    #' @param sheet sheet
    #' @returns The `wbWorkbook` object
    remove_filter = function(sheet = current_sheet()) {
      workbook_remove_filter(self, private, sheet)
    },

    ### grid lines ----

    #' @description grid lines
    #' @param sheet sheet
    #' @param show show
    #' @param print print
    #' @returns The `wbWorkbook` object
    grid_lines = function(sheet = current_sheet(), show = FALSE, print = show) {
      workbook_grid_lines(
        self    = self,
        private = private,
        sheet   = sheet,
        show    = show,
        print   = print
      )
    },

    ### named region ----

    #' @description add a named region
    #' @param sheet sheet
    #' @param cols cols
    #' @param rows rows
    #' @param name name
    #' @param localSheet localSheet
    #' @param overwrite overwrite
    #' @param comment comment
    #' @param customMenu customMenu
    #' @param description description
    #' @param is_function function
    #' @param functionGroupId function group id
    #' @param help help
    #' @param hidden hidden
    #' @param localName localName
    #' @param publishToServer publish to server
    #' @param statusBar status bar
    #' @param vbProcedure wbProcedure
    #' @param workbookParameter workbookParameter
    #' @param xml xml
    #' @returns The `wbWorkbook` object
    add_named_region = function(
      sheet = current_sheet(),
      cols,
      rows,
      name,
      localSheet        = FALSE,
      overwrite         = FALSE,
      comment           = NULL,
      customMenu        = NULL,
      description       = NULL,
      is_function       = NULL,
      functionGroupId   = NULL,
      help              = NULL,
      hidden            = NULL,
      localName         = NULL,
      publishToServer   = NULL,
      statusBar         = NULL,
      vbProcedure       = NULL,
      workbookParameter = NULL,
      xml               = NULL
    ) {
      workbook_add_named_region(
        self              = self,
        private           = private,
        sheet             = sheet,
        cols              = cols,
        rows              = rows,
        name              = name,
        localSheet        = localSheet,
        overwrite         = overwrite,
        comment           = comment,
        customMenu        = customMenu,
        description       = description,
        is_function       = is_function,
        functionGroupId   = functionGroupId,
        help              = help,
        hidden            = hidden,
        localName         = localName,
        publishToServer   = publishToServer,
        statusBar         = statusBar,
        vbProcedure       = vbProcedure,
        workbookParameter = workbookParameter,
        xml               = xml
      )
    },

    #' @description remove a named region
    #' @param sheet sheet
    #' @param name name
    #' @returns The `wbWorkbook` object
    remove_named_region = function(sheet = current_sheet(), name = NULL) {
      workbook_remove_named_region(
        self    = self,
        private = private,
        sheet   = sheet,
        name    = name
      )
    },

    ### order ----

    #' @description set worksheet order
    #' @param sheets sheets
    #' @return The `wbWorkbook` object
    set_order = function(sheets) {
      workbook_set_order(self, private, sheets)
    },

    ### sheet visibility ----

    #' @description Get sheet visibility
    #' @returns Returns sheet visibility
    get_sheet_visibility = function() {
      workbook_get_sheet_visibility(self, private)
    },

    #' @description Set sheet visibility
    #' @param value value
    #' @param sheet sheet
    #' @returns The `wbWorkbook` object
    set_sheet_visibility = function(sheet = current_sheet(), value) {
      workbook_set_sheet_visibility(
        self    = self,
        private = private,
        sheet   = sheet,
        value   = value
      )
    },

    ### page breaks ----

    #' @description Add a page break
    #' @param sheet sheet
    #' @param row row
    #' @param col col
    #' @returns The `wbWorkbook` object
    add_page_break = function(sheet = current_sheet(), row = NULL, col = NULL) {
      workbook_add_page_breaks(
        self    = self,
        private = private,
        sheet   = sheet,
        row     = row,
        col     = col
      )
    },

    ### clean sheet----

    #' @description clean sheet (remove all values)
    #' @param sheet sheet
    #' @param numbers remove all numbers
    #' @param characters remove all characters
    #' @param styles remove all styles
    #' @param merged_cells remove all merged_cells
    #' @return The `wbWorksheetObject`, invisibly
    clean_sheet = function(
      sheet        = current_sheet(),
      numbers      = TRUE,
      characters   = TRUE,
      styles       = TRUE,
      merged_cells = TRUE
    ) {
      workbook_clean_sheet(
        self         = self,
        private      = private,
        sheet        = sheet,
        numbers      = numbers,
        characters   = characters,
        styles       = styles,
        merged_cells = merged_cells
      )
    },

    ### styles ----

    #' @description create borders for cell region
    #' @param sheet a worksheet
    #' @param dims dimensions on the worksheet e.g. "A1", "A1:A5", "A1:H5"
    #' @param bottom_color,left_color,right_color,top_color,inner_hcolor,inner_vcolor a color, either something openxml knows or some RGB color
    #' @param left_border,right_border,top_border,bottom_border,inner_hgrid,inner_vgrid the border style, if NULL no border is drawn. See create_border for possible border styles
    #' @seealso create_border
    #' @examples
    #'
    #' wb <- wb_workbook()
    #' wb$add_worksheet("S1")$add_data("S1", mtcars)
    #' wb$add_border(1, dims = "A1:K1",
    #'  left_border = NULL, right_border = NULL,
    #'  top_border = NULL, bottom_border = "double")
    #' wb$add_border(1, dims = "A5",
    #'  left_border = "dotted", right_border = "dotted",
    #'  top_border = "hair", bottom_border = "thick")
    #' wb$add_border(1, dims = "C2:C5")
    #' wb$add_border(1, dims = "G2:H3")
    #' wb$add_border(1, dims = "G12:H13",
    #'  left_color = wb_colour(hex = "FF9400D3"), right_color = wb_colour(hex = "FF4B0082"),
    #'  top_color = wb_colour(hex = "FF0000FF"), bottom_color = wb_colour(hex = "FF00FF00"))
    #' wb$add_border(1, dims = "A20:C23")
    #' wb$add_border(1, dims = "B12:D14",
    #'  left_color = wb_colour(hex = "FFFFFF00"), right_color = wb_colour(hex = "FFFF7F00"),
    #'  bottom_color = wb_colour(hex = "FFFF0000"))
    #' wb$add_border(1, dims = "D28:E28")
    #' # if (interactive()) wb$open()
    #'
    #' wb <- wb_workbook()
    #' wb$add_worksheet("S1")$add_data("S1", mtcars)
    #' wb$add_border(1, dims = "A2:K33", inner_vgrid = "thin", inner_vcolor = c(rgb="FF808080"))
    #' @return The `wbWorksheetObject`, invisibly
    add_border = function(
    sheet         = current_sheet(),
    dims          = "A1",
    bottom_color  = wb_colour(hex = "FF000000"),
    left_color    = wb_colour(hex = "FF000000"),
    right_color   = wb_colour(hex = "FF000000"),
    top_color     = wb_colour(hex = "FF000000"),
    bottom_border = "thin",
    left_border   = "thin",
    right_border  = "thin",
    top_border    = "thin",
    inner_hgrid   = NULL,
    inner_hcolor  = NULL,
    inner_vgrid   = NULL,
    inner_vcolor  = NULL
    ) {
      workbook_add_border(
        self          = self,
        private       = private,
        sheet         = sheet,
        dims          = dims,
        bottom_color  = bottom_color,
        left_color    = left_color,
        right_color   = right_color,
        top_color     = top_color,
        bottom_border = bottom_border,
        left_border   = left_border,
        right_border  = right_border,
        top_border    = top_border,
        inner_hgrid   = inner_hgrid,
        inner_hcolor  = inner_hcolor,
        inner_vgrid   = inner_vgrid,
        inner_vcolor  = inner_vcolor
      )
    },

    #' @description provide simple fill function
    #' @param sheet the worksheet
    #' @param dims the cell range
    #' @param color the colors to apply, e.g. yellow: wb_colour(hex = "FFFFFF00")
    #' @param pattern various default "none" but others are possible:
    #'  "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal",
    #'  "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis",
    #'  "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid",
    #'  "lightTrellis", "gray125", "gray0625"
    #' @param gradient_fill a gradient fill xml pattern.
    #' @param every_nth_col which col should be filled
    #' @param every_nth_row which row should be filled
    #' @examples
    #'  # example from the gradient fill manual page
    #'  gradient_fill <- "<gradientFill degree=\"90\">
    #'    <stop position=\"0\"><color rgb=\"FF92D050\"/></stop>
    #'    <stop position=\"1\"><color rgb=\"FF0070C0\"/></stop>
    #'   </gradientFill>"
    #' @return The `wbWorksheetObject`, invisibly
    add_fill = function(
      sheet         = current_sheet(),
      dims          = "A1",
      color         = wb_colour(hex = "FFFFFF00"),
      pattern       = "solid",
      gradient_fill = "",
      every_nth_col = 1,
      every_nth_row = 1
    ) {
      workbook_add_fill(
        self          = self,
        private       = private,
        sheet         = sheet,
        dims          = dims,
        color         = color,
        pattern       = pattern,
        gradient_fill = gradient_fill,
        every_nth_col = every_nth_col,
        every_nth_row = every_nth_row
      )
    },

    #' @description provide simple font function
    #' @param sheet the worksheet
    #' @param dims the cell range
    #' @param name font name: default "Calibri"
    #' @param color rgb color: default "FF000000"
    #' @param size font size: default "11",
    #' @param bold bold
    #' @param italic italic
    #' @param outline outline
    #' @param strike strike
    #' @param underline underline
    #' @param family font family
    #' @param charset charset
    #' @param condense condense
    #' @param scheme font scheme
    #' @param shadow shadow
    #' @param extend extend
    #' @param vertAlign vertical alignment
    #' @examples
    #'  wb <- wb_workbook()$add_worksheet("S1")$add_data("S1", mtcars)
    #'  wb$add_font("S1", "A1:K1", name = "Arial", color = wb_colour(theme = "4"))
    #' @return The `wbWorksheetObject`, invisibly
    add_font = function(
      sheet     = current_sheet(),
      dims      = "A1",
      name      = "Calibri",
      color     = wb_colour(hex = "FF000000"),
      size      = "11",
      bold      = "",
      italic    = "",
      outline   = "",
      strike    = "",
      underline = "",
      # fine tuning
      charset   = "",
      condense  = "",
      extend    = "",
      family    = "",
      scheme    = "",
      shadow    = "",
      vertAlign = ""
    ) {
      workbook_add_font(
        self      = self,
        private   = private,
        sheet     = sheet,
        dims      = dims,
        name      = name,
        color     = color,
        size      = size,
        bold      = bold,
        italic    = italic,
        outline   = outline,
        strike    = strike,
        underline = underline,
        charset   = charset,
        condense  = condense,
        extend    = extend,
        family    = family,
        scheme    = scheme,
        shadow    = shadow,
        vertAlign = vertAlign
      )
    },

    #' @description provide simple number format function
    #' @param sheet the worksheet
    #' @param dims the cell range
    #' @param numfmt number format id or a character of the format
    #' @examples
    #'  wb <- wb_workbook()$add_worksheet("S1")$add_data("S1", mtcars)
    #'  wb$add_numfmt("S1", "A1:A33", numfmt = 1)
    #' @return The `wbWorksheetObject`, invisibly
    add_numfmt = function(
      sheet = current_sheet(),
      dims  = "A1",
      numfmt
    ) {
      workbook_add_numfmt(
        self    = self,
        private = private,
        sheet   = sheet,
        dims    = dims,
        numfm   = numfmt
      )
    },

    #' @description provide simple cell style format function
    #' @param sheet the worksheet
    #' @param dims the cell range
    #' @param extLst extension list something like `<extLst>...</extLst>`
    #' @param hidden logical cell is hidden
    #' @param horizontal align content horizontal ('left', 'center', 'right')
    #' @param indent logical indent content
    #' @param justifyLastLine logical justify last line
    #' @param locked logical cell is locked
    #' @param pivotButton unknown
    #' @param quotePrefix unknown
    #' @param readingOrder reading order left to right
    #' @param relativeIndent relative indentation
    #' @param shrinkToFit logical shrink to fit
    #' @param textRotation degrees of text rotation
    #' @param vertical vertical alignment of content ('top', 'center', 'bottom')
    #' @param wrapText wrap text in cell
    # alignments
    #' @param applyAlignment logical apply alignment
    #' @param applyBorder logical apply border
    #' @param applyFill logical apply fill
    #' @param applyFont logical apply font
    #' @param applyNumberFormat logical apply number format
    #' @param applyProtection logical apply protection
    # ids
    #' @param borderId border ID to apply
    #' @param fillId fill ID to apply
    #' @param fontId font ID to apply
    #' @param numFmtId number format ID to apply
    #' @param xfId xf ID to apply
    #' @examples
    #'  wb <- wb_workbook()$add_worksheet("S1")$add_data("S1", mtcars)
    #'  wb$add_cell_style("S1", "A1:K1",
    #'                    textRotation = "45",
    #'                    horizontal = "center",
    #'                    vertical = "center",
    #'                    wrapText = "1")
    #' @return The `wbWorksheetObject`, invisibly
    add_cell_style = function(
      sheet             = current_sheet(),
      dims              = "A1",
      applyAlignment    = NULL,
      applyBorder       = NULL,
      applyFill         = NULL,
      applyFont         = NULL,
      applyNumberFormat = NULL,
      applyProtection   = NULL,
      borderId          = NULL,
      extLst            = NULL,
      fillId            = NULL,
      fontId            = NULL,
      hidden            = NULL,
      horizontal        = NULL,
      indent            = NULL,
      justifyLastLine   = NULL,
      locked            = NULL,
      numFmtId          = NULL,
      pivotButton       = NULL,
      quotePrefix       = NULL,
      readingOrder      = NULL,
      relativeIndent    = NULL,
      shrinkToFit       = NULL,
      textRotation      = NULL,
      vertical          = NULL,
      wrapText          = NULL,
      xfId              = NULL
    ) {
      workbook_add_cell_style(
        self              = self,
        private           = private,
        sheet             = sheet,
        dims              = dims,
        applyAlignment    = applyAlignment,
        applyBorder       = applyBorder,
        applyFill         = applyFill,
        applyFont         = applyFont,
        applyNumberFormat = applyNumberFormat,
        applyProtection   = applyProtection,
        borderId          = borderId,
        extLst            = extLst,
        fillId            = fillId,
        fontId            = fontId,
        hidden            = hidden,
        horizontal        = horizontal,
        indent            = indent,
        justifyLastLine   = justifyLastLine,
        locked            = locked,
        numFmtId          = numFmtId,
        pivotButton       = pivotButton,
        quotePrefix       = quotePrefix,
        readingOrder      = readingOrder,
        relativeIndent    = relativeIndent,
        shrinkToFit       = shrinkToFit,
        textRotation      = textRotation,
        vertical          = vertical,
        wrapText          = wrapText,
        xfId              = xfId
      )
    },

    #' @description get sheet style
    #' @param sheet sheet
    #' @param dims dims
    #' @returns a character vector of cell styles
    get_cell_style = function(sheet = current_sheet(), dims) {
      workbook_get_cell_style(
        self    = self,
        private = private,
        sheet   = sheet,
        dims    = dims
      )
    },

    #' @description set sheet style
    #' @param sheet sheet
    #' @param dims dims
    #' @param style style
    #' @return The `wbWorksheetObject`, invisibly
    set_cell_style = function(sheet = current_sheet(), dims, style) {
      workbook_set_cell_style(
        self    = self,
        private = private,
        sheet   = sheet,
        dims    = dims,
        style   = style
      )
    },

    #' @description clone style from one sheet to another
    #' @param from the worksheet you are cloning
    #' @param to the worksheet the style is applied to
    clone_sheet_style = function(from = current_sheet(), to) {
      workbook_clone_sheet_style(self, private, from, to)
    },

    #' @description apply sparkline to worksheet
    #' @param sheet the worksheet you are using
    #' @param sparklines sparkline created by `create_sparkline()`
    add_sparklines = function(sheet = current_sheet(), sparklines) {
      workbook_add_sparklines(self, private, sheet, sparklines)
    }
  ),

  ## private ----

  # any functions that are not present elsewhere or are non-exported internal
  # functions that are used to make assignments
  private = list(
    ### fields ----
    current_sheet = 0L,

    # original sheet name values
    original_sheet_names = character(),

    ### methods ----
    deep_clone = function(name, value) {
      # Deep cloning method for workbooks.  This method also accesses
      # `$clone(deep = TRUE)` methods for `R6` fields.
      if (R6::is.R6(value)) {
        value <- value$clone(deep = TRUE)
      } else if (is.list(value)) {
        # specifically targetting fields like `worksheets`
        for (i in wapply(value, R6::is.R6)) {
          value[[i]] <- value[[i]]$clone(deep = TRUE)
        }
      }

      value
    },

    pappend = function(field, value = NULL) {
      # private append
      private[[field]] <- c(private[[field]], value)
    },

    validate_new_sheet = function(sheet) {
      # returns nothing, throws error if there's a problem.
      if (length(sheet) != 1) {
        stop("sheet name must be length 1")
      }

      if (is_waiver(sheet)) {
        # should be safe
        return()
      }

      if (is.na(sheet)) {
        stop("sheet cannot be NA")
      }

      if (is.numeric(sheet)) {
        if (!is_integer_ish(sheet)) {
          stop("If sheet is numeric it must be an integer")
        }

        if (sheet <= 0) {
          stop("if sheet is an integer it must be a positive number")
        }

        if (sheet <= length(self$sheet_names)) {
          stop("there is already a sheet at index ", sheet)
        }

        return()
      }

      sheet <- as.character(sheet)
      if (has_illegal_chars(sheet)) {
        stop("illegal characters found in sheet. Please remove. See ?openxlsx::clean_worksheet_name")
      }

      if (!nzchar(sheet)) {
        stop("sheet name must contain at least 1 character")
      }

      if (nchar(sheet) > 31) {
        stop("sheet names must be <= 31 chars")
      }

      if (tolower(sheet) %in% self$sheet_names) {
        stop("a sheet with name '", sheet, '"already exists"')
      }
    },

    set_current_sheet = function(sheet_index) {
      stopifnot(is_integer_ish(sheet_index), length(sheet_index) == 1)
      private$current_sheet <- sheet_index
    },

    get_sheet_index = function(sheet) {
      # Get/validate `sheet` and set as the current sheet
      if (is_waiver(sheet)) {
        # waivers shouldn't need additional validation
        switch(
          sheet,
          current_sheet = NULL,
          next_sheet = {
            private$current_sheet <- length(self$sheet_names) + 1L
          },
          stop("not a valid waiver: ", sheet)
        )
        return(private$current_sheet)
      }

      # returns the sheet index, or NA
      if (is.null(self$sheet_names)) {
        warning("Workbook has no sheets")
        return(NA_integer_)
      }

      if (is.character(sheet)) {
        sheet <- tolower(sheet)
        m1 <- match(sheet, tolower(self$sheet_names))
        m2 <- match(sheet, tolower(private$original_sheet_names))

        bad <- is.na(m1) & is.na(m2)

        if (any(bad)) {
          stop("Sheet name(s) not found: ", toString(sheet[bad]))
        }

        # need the vectorized
        sheet <- ifelse(is.na(m1), m2, m1)
      } else {
        sheet <- as.integer(sheet)
        bad <- which(sheet > length(self$sheet_names) | sheet < 1)

        if (length(bad)) {
          stop("Invalid sheet position(s): ", toString(sheet[bad]))
          # sheet[bad] <- NA_integer_
        }
      }

      private$current_sheet <- sheet
      sheet
    },

    get_sheet_name = function(sheet) {
      self$sheet_names[private$get_sheet_index(sheet)]
    },

    set_single_sheet_name = function(pos, clean, raw) {
      pos <- as.integer(pos)
      stopifnot(
        length(pos)   == 1, !is.na(pos),
        length(clean) == 1, !is.na(clean),
        length(raw)   == 1, !is.na(raw)
      )
      self$sheet_names[self$sheetOrder[pos]] <- clean
      private$original_sheet_names[self$sheetOrder[pos]] <- raw
    },

    append_sheet_field = function(sheet = current_sheet(), field, value = NULL) {
      # if using this we should consider adding a method into the wbWorksheet
      # object.  wbWorksheet$append() is currently public. _Currently_.
      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$append(field, value)
      invisible(self)
    },

    append_workbook_field = function(field, value = NULL) {
      self$workbook[[field]] <- c(self$workbook[[field]], value)
      invisible(self)
    },

    append_sheet_rels = function(sheet = current_sheet(), value = NULL) {
      sheet <- private$get_sheet_index(sheet)
      self$worksheets_rels[[sheet]] <- c(self$worksheets_rels[[sheet]], value)
      invisible(self)
    },

    generate_base_core = function() {

      self$core <-
        paste_c(
          # base
          paste(
            '<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"',
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"',
            'xmlns:dc="http://purl.org/dc/elements/1.1/"',
            'xmlns:dcterms="http://purl.org/dc/terms/"',
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
          ),

          # non-optional
          sprintf("<dc:creator>%s</dc:creator>",                                     paste(self$creator, collapse = ";")),
          sprintf("<cp:lastModifiedBy>%s</cp:lastModifiedBy>",                       paste(self$creator, collapse = ";")),
          sprintf('<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>', format(self$datetimeCreated, "%Y-%m-%dT%H:%M:%SZ")),

          # optional
          if (!is.null(self$title))    sprintf("<dc:title>%s</dc:title>",       replace_legal_chars(self$title)),
          if (!is.null(self$subject))  sprintf("<dc:subject>%s</dc:subject>",   replace_legal_chars(self$subject)),
          if (!is.null(self$category)) sprintf("<cp:category>%s</cp:category>", replace_legal_chars(self$category)),

          # end
          "</coreProperties>",
          collapse = "",
          unlist = TRUE
        )

      invisible(self)
    },

    get_worksheet = function(sheet) {
      self$worksheets[[private$get_sheet_index(sheet)]]
    },

    # this may ahve been removes
    updateSharedStrings = function(uNewStr) {
      ## Function will return named list of references to new strings
      uStr <- uNewStr[which(!uNewStr %in% self$sharedStrings)]
      uCount <- attr(self$sharedStrings, "uniqueCount")
      self$append("sharedStrings", uStr)

      attr(self$sharedStrings, "uniqueCount") <- uCount + length(uStr)
      invisible(self)
    },

    writeDrawingVML = function(dir, dir_rel) {

      # not sure if comments and vml are the same length
      counter <- max(length(self$comments), length(self$vml))

      # beg vml loop
      for (i in seq_len(counter)) {
        id <- 1025

        vml_ext <- NULL

        ## get additional vml
        if (!is.null(unlist(self$vml[i]))) {
          if (length(self$vml[[i]])) {
            vml_ext <- c(vml_ext, getXMLPtr1con(read_xml(self$vml[[i]])))
          }
        }

        vml_comment <- NULL

        ## get comment vml
        if (!is.null(unlist(self$comments[i]))) {
          cd <- unapply(self$comments[[i]], "[[", "clientData")
          nComments <- length(cd)

          vml_comment <- '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'

          for (j in seq_len(nComments)) {
            id <- id + 1L
            vml_comment <- c(
              vml_comment, genBaseShapeVML(cd[j], id)
            )
          }
        }

        vml_xml <- c(vml_ext, vml_comment)


        ## create output only if vml_comment != NULL
        if (!is.null(vml_xml)) {

          # keep only the first o:shapelayout
          vml_xml <- xml_node(vml_xml)
          oshapelayout <- which(xml_node_name(vml_xml) == "o:shapelayout")
          sel <- which(!seq_along(vml_xml) %in% oshapelayout[-1])

          ## create vml for output
          vml_xml <-  xml_node_create(
            xml_name = "xml",
            xml_attributes = c(
              `xmlns:v` = "urn:schemas-microsoft-com:vml",
              `xmlns:o` = "urn:schemas-microsoft-com:office:office",
              `xmlns:x` = "urn:schemas-microsoft-com:office:excel"
            ),
            xml_children = vml_xml[sel]
          )

          ## write vml output
          write_file(
            head = '',
            body = pxml(vml_xml),
            tail = '',
            fl = file.path(dir, sprintf("vmlDrawing%s.vml", i))
          )

          ## vml drawing
          if (length(self$vml_rels[[i]])) {
            write_file(
              head = '',
              body = pxml(self$vml_rels[[i]]),
              tail = '',
              fl = file.path(dir_rel, stri_join("vmlDrawing", i, ".vml.rels"))
            )
          }
        }

      } # end vml loop

      invisible(self)
    },

    writeSheetDataXML = function(
      ct,
      xldrawingsDir,
      xldrawingsRelsDir,
      xlchartsDir,
      xlchartsRelsDir,
      xlworksheetsDir,
      xlworksheetsRelsDir
    ) {

      ## write charts
      if (NROW(self$charts) && any(self$charts != "")) {

        if (!file.exists(xlchartsDir)) {
          dir.create(xlchartsDir, recursive = TRUE)
          if (any(self$charts$rels != "") || any(self$charts$relsEx != ""))
            dir.create(xlchartsRelsDir, recursive = TRUE)
        }

        for (crt in seq_len(nrow(self$charts))) {

          if (self$charts$chart[crt] != "") {
            ct <- c(ct, sprintf('<Override PartName="/xl/charts/chart%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', crt))

            write_file(
              body = self$charts$chart[crt],
              fl = file.path(xlchartsDir, stri_join("chart", crt, ".xml"))
            )
          }

          if (self$charts$chartEx[crt] != "") {
            ct <- c(ct, sprintf('<Override PartName="/xl/charts/chartEx%s.xml" ContentType="application/vnd.ms-office.chartex+xml"/>', crt))

            write_file(
              body = self$charts$chartEx[crt],
              fl = file.path(xlchartsDir, stri_join("chartEx", crt, ".xml"))
            )
          }

          if (self$charts$colors[crt] != "") {
            ct <- c(ct, sprintf('<Override PartName="/xl/charts/colors%s.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>', crt))

            write_file(
              body = self$charts$colors[crt],
              fl = file.path(xlchartsDir, stri_join("colors", crt, ".xml"))
            )
          }

          if (self$charts$style[crt] != "") {
            ct <- c(ct, sprintf('<Override PartName="/xl/charts/style%s.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>', crt))

            write_file(
              body = self$charts$style[crt],
              fl = file.path(xlchartsDir, stri_join("style", crt, ".xml"))
            )
          }

          if (self$charts$rels[crt] != "") {
            write_file(
              body = self$charts$rels[crt],
              fl = file.path(xlchartsRelsDir, stri_join("chart", crt, ".xml.rels"))
            )
          }

          if (self$charts$relsEx[crt] != "") {
            write_file(
              body = self$charts$relsEx[crt],
              fl = file.path(xlchartsRelsDir, stri_join("chartEx", crt, ".xml.rels"))
            )
          }
        }

      }

      ## write drawings

      nDrawings <- length(self$drawings)

      for (i in seq_len(nDrawings)) {

        ## Write drawing i (will always exist) skip those that are empty
        if (!all(self$drawings[[i]] == "")) {
          write_file(
            head = '',
            body = pxml(self$drawings[[i]]),
            tail = '',
            fl = file.path(xldrawingsDir, stri_join("drawing", i, ".xml"))
          )
          if (!all(self$drawings_rels[[i]] == "")) {
            write_file(
              head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
              body = pxml(self$drawings_rels[[i]]),
              tail = '</Relationships>',
              fl = file.path(xldrawingsRelsDir, stri_join("drawing", i, ".xml.rels"))
            )
          }

          drawing_type <- xml_node_name(self$drawings[[i]])
          if (drawing_type == "xdr:wsDr") {
            ct_drawing <- sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', i)
          } else if (drawing_type == "c:userShapes") {
            ct_drawing <- sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"/>', i)
          }

          ct <- c(ct, ct_drawing)

        }

      }

      ## write worksheets

      # TODO just seq_along()
      nSheets <- length(self$worksheets)

      for (i in seq_len(nSheets)) {

        if (self$isChartSheet[i]) {
          chartSheetDir <- file.path(dirname(xlworksheetsDir), "chartsheets")
          chartSheetRelsDir <-
            file.path(dirname(xlworksheetsDir), "chartsheets", "_rels")

          if (!file.exists(chartSheetDir)) {
            dir.create(chartSheetDir, recursive = FALSE)
            dir.create(chartSheetRelsDir, recursive = FALSE)
          }

          write_file(
            body = self$worksheets[[i]]$get_prior_sheet_data(),
            fl = file.path(chartSheetDir, stri_join("sheet", i, ".xml"))
          )

          if (length(self$worksheets_rels[[i]])) {
            write_file(
              head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
              body = pxml(self$worksheets_rels[[i]]),
              tail = "</Relationships>",
              fl = file.path(chartSheetRelsDir, sprintf("sheet%s.xml.rels", i))
            )
          }
        } else {
          ## Write worksheets
          ws <- self$worksheets[[i]]
          hasHL <- length(ws$hyperlinks) > 0

          prior <- ws$get_prior_sheet_data()
          post <- ws$get_post_sheet_data()

          cc <- ws$sheet_data$cc


          if (!is.null(cc)) {
            cc$r <- stri_join(cc$c_r, cc$row_r)
            # prepare data for output

            # there can be files, where row_attr is incomplete because a row
            # is lacking any attributes (presumably was added before saving)
            # still row_attr is what we want!

            rows_attr <- ws$sheet_data$row_attr
            ws$sheet_data$row_attr <- rows_attr[order(as.numeric(rows_attr[, "r"])), ]

            cc_rows <- ws$sheet_data$row_attr$r
            cc_out <- cc[cc$row_r %in% cc_rows, c("row_r", "c_r",  "r", "v", "c_t", "c_s", "c_cm", "c_ph", "c_vm", "f", "f_t", "f_ref", "f_ca", "f_si", "is")]

            ws$sheet_data$cc_out <- cc_out[order(as.integer(cc_out[, "row_r"]), col2int(cc_out[, "c_r"])), ]
          } else {
            ws$sheet_data$row_attr <- NULL
            ws$sheet_data$cc_out <- NULL
          }

          # row_attr <- ws$sheet_data$row_attr
          # nam_at <- names(row_attr)
          # wanted <- as.character(seq(min(as.numeric(nam_at)),
          #                            max(as.numeric(nam_at))))
          # empty_row_attr <- wanted[!wanted %in% nam_at]
          # # add empty list
          # if (!identical(empty_row_attr, character()))
          #   row_attr[[empty_row_attr]] <- list()
          # # restore order
          # ws$sheet_data$row_attr <- row_attr[wanted]

          # create entire sheet prior to writing it
          sheet_xml <- write_worksheet(
            prior = prior,
            post = post,
            sheet_data = ws$sheet_data
          )
          write_xmlPtr(doc = sheet_xml, fl = file.path(xlworksheetsDir, sprintf("sheet%s.xml", i)))

          ## write worksheet rels
          if (length(self$worksheets_rels[[i]])) {
            ws_rels <- self$worksheets_rels[[i]]
            if (hasHL) {
              h_inds <- stri_join(seq_along(self$worksheets[[i]]$hyperlinks), "h")
              ws_rels <-
                c(ws_rels, unlist(
                  lapply(seq_along(h_inds), function(j) {
                    self$worksheets[[i]]$hyperlinks[[j]]$to_target_xml(h_inds[j])
                  })
                ))
            }

            ## Check if any tables were deleted - remove these from rels
            # TODO a relship manager should take care of this
            tabs <- self$tables[self$tables$tab_act == 1, ]
            if (NROW(tabs)) {
              table_inds <- grep("tables/table[0-9]+.xml", ws_rels)

              relship <- rbindlist(xml_attr(ws_rels, "Relationship"))
              if (ncol(relship) && nrow(relship)) {
                relship$typ <- basename(relship$Type)
                relship$tid <- as.numeric(gsub("\\D+", "", relship$Target))

                relship$typ <- relship$tid <- NULL
                if (is.null(relship$TargetMode)) relship$TargetMode <- ""
                ws_rels <- df_to_xml("Relationship", df_col = relship[c("Id", "Type", "Target", "TargetMode")])
              }
            }


            # if (i < 3)
            write_file(
              head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
              body = pxml(ws_rels),
              tail = "</Relationships>",
              fl = file.path(xlworksheetsRelsDir, sprintf("sheet%s.xml.rels", i))
            )
          }
        } ## end of isChartSheet[i]
      } ## end of loop through nSheets

      return(ct)
    },

    # old add_named_region()
    create_named_region = function(
      ref1,
      ref2,
      name,
      sheet = current_sheet(),
      localSheetId      = NULL,
      comment           = NULL,
      customMenu        = NULL,
      description       = NULL,
      is_function       = NULL,
      functionGroupId   = NULL,
      help              = NULL,
      hidden            = NULL,
      localName         = NULL,
      publishToServer   = NULL,
      statusBar         = NULL,
      vbProcedure       = NULL,
      workbookParameter = NULL,
      xml               = NULL
    ) {
      name <- replace_legal_chars(name)

      # special names

      ## print
      # _xlnm .Print_Area
      # _xlnm .Print_Titles

      ## filters
      # _xlnm .Criteria
      # _xlnm ._FilterDatabase
      # _xlnm .Extract

      ## misc
      # _xlnm .Consolidate_Area
      # _xlnm .Database
      # _xlnm .Sheet_Title

      named_region <- c(
        comment          = comment,
        customMenu        = customMenu,
        description       = description,
        `function`        = is_function,
        functionGroupId   = functionGroupId,
        help              = help,
        hidden            = hidden,
        localName         = localName,
        localSheetId      = localSheetId,
        name              = name,
        publishToServer   = publishToServer,
        statusBar         = statusBar,
        vbProcedure       = vbProcedure,
        workbookParameter = workbookParameter,
        xml               = xml
      )

      xml <- xml_node_create(
        "definedName",
        xml_children = sprintf("\'%s\'!%s:%s", sheet, ref1, ref2),
        xml_attributes = named_region
      )

      private$append_workbook_field("definedNames", xml)
    },

    get_sheet_id = function(type = c("rId", "sheetId"), i = NULL) {
      pattern <-
        switch(
          match.arg(type),
          sheetId = '(?<=sheetId=")[0-9]+',
          rId = '(?<= r:id="rId)[0-9]+'
        )

      i <- i %||% seq_along(self$workbook$sheets)
      as.integer(unlist(reg_match0(self$workbook$sheets[i], pattern)))
    },

    get_sheet_id_max = function(i = NULL) {
      max(private$get_sheet_id(type = "sheetId", i = i), 0L, na.rm = TRUE) + 1L
    },

    do_conditional_formatting = function(
      sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      dxfId,
      formula,
      type,
      values,
      params
    ) {
      # TODO consider defaults for logicals
      # TODO rename: setConditionFormatting?  Or addConditionalFormatting
      # TODO can this be moved to the sheet data?
      sheet <- private$get_sheet_index(sheet)
      sqref <- stri_join(
        get_cell_refs(data.frame(x = c(startRow, endRow), y = c(startCol, endCol))),
        collapse = ":"
      )

      ## Increment priority of conditional formatting rule
      for (i in rev(seq_along(self$worksheets[[sheet]]$conditionalFormatting))) {
        priority <- reg_match0(
          self$worksheets[[sheet]]$conditionalFormatting[[i]],
          '(?<=priority=")[0-9]+'
        )
        priority_new <- as.integer(priority) + 1L
        priority_pattern <- sprintf('priority="%s"', priority)
        priority_new <- sprintf('priority="%s"', priority_new)

        ## now replace
        self$worksheets[[sheet]]$conditionalFormatting[[i]] <- gsub(
          priority_pattern,
          priority_new,
          self$worksheets[[sheet]]$conditionalFormatting[[i]],
          fixed = TRUE
        )
      }

      nms <- c(names(self$worksheets[[sheet]]$conditionalFormatting), sqref)
      dxfId <- max(dxfId, 0L)

      # big switch statement
      cfRule <- switch(
        type,

        ## colourScale ----
        colorScale = cf_create_colorscale(formula, values),

        ## dataBar ----
        dataBar = cf_create_databar(self$worksheets[[sheet]]$extLst, formula, params, sqref, values),

        ## expression ----
        expression = cf_create_expression(dxfId, formula),

        ## duplicatedValues ----
        duplicatedValues = cf_create_duplicated_values(dxfId),

        ## containsText ----
        containsText = cf_create_contains_text(dxfId, sqref, values),

        ## notContainsText ----
        notContainsText = cf_create_not_contains_text(dxfId, sqref, values),

        ## beginsWith ----
        beginsWith = cf_begins_with(dxfId, sqref, values),

        ## endsWith ----
        endsWith = cf_ends_with(dxfId, sqref, values),

        ## between ----
        between = cf_between(dxfId, formula),

        ## topN ----
        topN = cf_top_n(dxfId, values),

        ## bottomN ----
        bottomN = cf_bottom_n(dxfId, values),

        # do we have a match.arg() anywhere or will it just be showned in this switch()?
        stop("type `", type, "` is not a valid formatting rule")
      )

      # dataBar needs additional extLst
      if (!is.null(attr(cfRule, "extLst")))
        self$worksheets[[sheet]]$extLst <- read_xml(attr(cfRule, "extLst"), pointer = FALSE)

      private$append_sheet_field(sheet, "conditionalFormatting", read_xml(cfRule, pointer = FALSE))
      names(self$worksheets[[sheet]]$conditionalFormatting) <- nms
      invisible(self)
    },

    preSaveCleanUp = function() {
      # TODO consider name self$workbook_validate() ?

      ## Steps
      # Order workbook.xml.rels:
      #   sheets -> style -> theme -> sharedStrings -> persons -> tables -> calcChain
      # Assign workbook.xml.rels children rIds, seq_along(workbook.xml.rels)
      # Assign workbook$sheets rIds nSheets
      #
      ## drawings will always be r:id1 on worksheet
      ## tables will always have r:id equal to table xml file number tables/table(i).xml

      ## Every worksheet has a drawingXML as r:id 1
      ## Tables from r:id 2
      ## HyperLinks from nTables+3 to nTables+3+nHyperLinks-1
      ## vmlDrawing to have rId

      # browser()

      sheetRIds <- private$get_sheet_id("rId")
      nSheets   <- length(sheetRIds)
      nExtRefs  <- length(self$externalLinks)
      nPivots   <- length(self$pivotDefinitions)

      ## add a worksheet if none added
      if (nSheets == 0) {
        warning(
          "Workbook does not contain any worksheets. A worksheet will be added.",
          call. = FALSE
        )
        self$add_worksheet("Sheet 1")
        nSheets <- 1L
      }

      ## get index of each child element for ordering
      sheetInds        <- grep("(worksheets|chartsheets)/sheet[0-9]+\\.xml", self$workbook.xml.rels)
      stylesInd        <- grep("styles\\.xml",                               self$workbook.xml.rels)
      themeInd         <- grep("theme/theme[0-9]+.xml",                      self$workbook.xml.rels)
      connectionsInd   <- grep("connections.xml",                            self$workbook.xml.rels)
      customXMLInd     <- grep("customXml/item[0-9]+.xml",                   self$workbook.xml.rels)
      extRefInds       <- grep("externalLinks/externalLink[0-9]+.xml",       self$workbook.xml.rels)
      sharedStringsInd <- grep("sharedStrings.xml",                          self$workbook.xml.rels)
      tableInds        <- grep("table[0-9]+.xml",                            self$workbook.xml.rels)
      personInds       <- grep("person.xml",                                 self$workbook.xml.rels)
      calcChainInd     <- grep("calcChain.xml",                              self$workbook.xml.rels)


      ## Reordering of workbook.xml.rels
      ## don't want to re-assign rIds for pivot tables or slicer caches
      pivotNode        <- grep("pivotCache/pivotCacheDefinition[0-9]+.xml", self$workbook.xml.rels, value = TRUE)
      slicerNode       <- grep("slicerCache[0-9]+.xml",                     self$workbook.xml.rels, value = TRUE)

      ## Reorder children of workbook.xml.rels
      self$workbook.xml.rels <-
        self$workbook.xml.rels[c(
          sheetInds,
          extRefInds,
          themeInd,
          connectionsInd,
          customXMLInd,
          stylesInd,
          sharedStringsInd,
          tableInds,
          personInds,
          calcChainInd
        )]

      ## Re assign rIds to children of workbook.xml.rels
      self$workbook.xml.rels <-
        unapply(
          seq_along(self$workbook.xml.rels),
          function(i) {
            gsub('(?<=Relationship Id="rId)[0-9]+',
                 i,
                 self$workbook.xml.rels[[i]],
                 perl = TRUE
            )
          }
        )

      self$append("workbook.xml.rels", c(pivotNode, slicerNode))

      if (length(self$metadata)) {
        self$append("workbook.xml.rels",
                    sprintf(
                      '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>',
                      1L + length(self$workbook.xml.rels)
                    )
        )
      }

      if (!is.null(self$vbaProject)) {
        self$append("workbook.xml.rels",
                    sprintf(
                      '<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>',
                      1L + length(self$workbook.xml.rels)
                    )
        )
      }

      ## Reassign rId to workbook sheet elements, (order sheets by sheetId first)
      self$workbook$sheets <-
        unapply(
          seq_along(self$workbook$sheets),
          function(i) {
            gsub('(?<= r:id="rId)[0-9]+', i, self$workbook$sheets[[i]], perl = TRUE)
          }
        )

      ## re-order worksheets if need to
      if (any(self$sheetOrder != seq_len(nSheets))) {
        self$workbook$sheets <- self$workbook$sheets[self$sheetOrder]
      }


      ## re-assign tabSelected
      state <- rep.int("visible", nSheets)
      hidden <- grepl("hidden", self$workbook$sheets)
      state[hidden] <- "hidden"
      visible_sheet_index <- which(!hidden)[1] # first visible

      if (is.null(self$workbook$bookViews))
        self$set_bookview(
          xWindow      = 0,
          yWindow      = 0,
          windowWidth  = 13125,
          windowHeight = 13125,
          firstSheet   = visible_sheet_index - 1L,
          activeTab    = visible_sheet_index - 1L
        )

      # Failsafe: hidden sheet can not be selected.
      self$worksheets[[visible_sheet_index]]$set_sheetview(tabSelected = TRUE)
      if (nSheets > 1) {
        for (i in setdiff(seq_len(nSheets), visible_sheet_index)) {
          self$worksheets[[i]]$set_sheetview(tabSelected = FALSE)
        }
      }

      ## update workbook r:id to match reordered workbook.xml.rels externalLink element
      if (length(extRefInds)) {
        newInds <- seq_along(extRefInds) + length(sheetInds)
        self$workbook$externalReferences <- stri_join(
          "<externalReferences>",
          stri_join(sprintf('<externalReference r:id=\"rId%s\"/>', newInds), collapse = ""),
          "</externalReferences>"
        )
      }

      invisible(self)
    },

    ## @description initialize cells in workbook
    ## @param sheet sheet
    ## @param dims dims
    ## @keywords internal
    do_cell_init = function(sheet = current_sheet(), dims) {

      sheet <- private$get_sheet_index(sheet)
      if (length(dims) == 1 && grepl(":|;", dims))
        dims <- dims_to_dataframe(dims, fill = TRUE)

      exp_cells <- unname(unlist(dims))
      got_cells <- self$worksheets[[sheet]]$sheet_data$cc$r

      # initialize cell
      if (!all(exp_cells %in% got_cells)) {

        init_cells <- NA
        missing_cells <- exp_cells[!exp_cells %in% got_cells]

        for (exp_cell in missing_cells) {
          self$add_data(
            x = init_cells,
            na.strings = NULL,
            colNames = FALSE,
            dims = exp_cell
          )
        }
      }

      invisible(self)
    }
  )
)


# helpers -----------------------------------------------------------------


file_copy_wb_save <- function(from, pattern, dir) {
  # specifically used within wbWoorkbook$save()

  stopifnot(
    dir.exists(dir),
    is.character(pattern),
    length(pattern) == 1L
  )

  if (length(from)) {
    file.copy(
      from = from,
      to = file.path(dir, sprintf(pattern, seq_along(from))),
      overwrite = TRUE,
      copy.date = TRUE
    )
  }
}

lcr <- function(var) {
  # quick function for specifying error message
  paste(var, "must have length 3 where elements correspond to positions: left, center, right.")
}


# TODO Does this need to be checked?  No sheet name can be NA right?
# res <- self$sheet_names[ind]; stopifnot(!anyNA(ind))

#' Get sheet name
#'
#' @param wb a [wbWorkbook] object
#' @param index Sheet name index
#' @return The sheet index
#' @export
wb_get_sheet_name <- function(wb, index = NULL) {
  index <- index %||% seq_along(wb$sheet_names)

  # index should be integer like
  stopifnot(is_integer_ish(index))

  n <- length(wb$sheet_names)

  if (any(index > n)) {
    stop("Invalid sheet index. Workbook ", n, " sheet(s)", call. = FALSE)
  }

  # keep index 0 as ""
  z <- vector("character", length(index))
  names(z) <- index
  z[index > 0] <- wb$sheet_names[index]
  z
}

worksheet_lock_properties <- function() {
  # provides a reference for the lock properties
  c(
    "selectLockedCells",
    "selectUnlockedCells",
    "formatCells",
    "formatColumns",
    "formatRows",
    "insertColumns",
    "insertRows",
    "insertHyperlinks",
    "deleteColumns",
    "deleteRows",
    "sort",
    "autoFilter",
    "pivotTables",
    "objects",
    "scenarios",
    NULL
  )
}
