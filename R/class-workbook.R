

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

    #' @field connections connections
    connections = NULL,

    #' @field Content_Types Content_Types
    Content_Types = genBaseContent_Type(),

    #' @field app app
    app = character(),

    #' @field core core
    core = character(),

    #' @field drawings drawings
    drawings = NULL,

    #' @field drawings_rels drawings_rels
    drawings_rels = NULL,

    #' @field drawings_vml drawings_vml
    drawings_vml = NULL,

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

    #' @field rowHeights rowHeights
    rowHeights = NULL,

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

    #' @field sheetOrder sheetOrder
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
      self$apps <- character()
      self$charts <- list()
      self$isChartSheet <- logical()

      self$connections <- NULL
      self$Content_Types <- genBaseContent_Type()
      self$core <-
        genBaseCore(
          creator = creator,
          title = title,
          subject = subject,
          category = category
        )
      self$comments <- list()
      self$threadComments <- list()


      self$drawings <- list()
      self$drawings_rels <- list()
      self$drawings_vml <- list()

      self$embeddings <- NULL
      self$externalLinks <- NULL
      self$externalLinksRels <- NULL

      self$headFoot <- NULL

      self$media <- list()

      self$persons <- NULL

      self$pivotTables <- NULL
      self$pivotTables.xml.rels <- NULL
      self$pivotDefinitions <- NULL
      self$pivotRecords <- NULL
      self$pivotDefinitionsRels <- NULL

      self$queryTables <- NULL
      self$rowHeights <- list()

      self$slicers <- NULL
      self$slicerCaches <- NULL

      self$sheet_names <- character()
      self$sheetOrder <- integer()

      self$sharedStrings <- list()
      attr(self$sharedStrings, "uniqueCount") <- 0

      self$styles_mgr <- style_mgr$new(self)
      self$styles_mgr$styles <- genBaseStyleSheet()

      self$tables <- NULL
      self$tables.xml.rels <- NULL
      self$theme <- NULL


      self$vbaProject <- NULL
      self$vml <- list()
      self$vml_rels <- list()

      self$creator <-
        creator %||%
        getOption("openxlsx2.creator") %||%
        # USERNAME may only be present for windows
        Sys.getenv("USERNAME", Sys.getenv("USER"))

      assert_class(self$creator,    "character")
      assert_class(title,           "character", or_null = TRUE)
      assert_class(subject,         "character", or_null = TRUE)
      assert_class(category,        "character", or_null = TRUE)
      assert_class(datetimeCreated, "POSIXt")

      stopifnot(
        length(title) <= 1L,
        length(category) <= 1L,
        length(datetimeCreated) == 1L
      )

      self$title           <- title
      self$subject         <- subject
      self$category        <- category
      self$datetimeCreated <- datetimeCreated
      private$generate_base_core()
      self
    },

    #' @description
    #' Append a field. This method is intended for internal use
    #' @param field A valid field name
    #' @param value A value for the field
    append = function(field, value) {
      self[[field]] <- c(self[[field]], value)
      self
    },

    #' @description
    #' Append to `self$workbook$sheets` This method is intended for internal use
    #' @param value A value for `self$workbook$sheets`
    append_sheets = function(value) {
      self$workbook$sheets <- c(self$workbook$sheets, value)
      self
    },

    #' @description validate sheet
    #' @param sheet A character sheet name or integer location
    #' @returns The integer position of the sheet
    validate_sheet = function(sheet) {

      # workbook has no sheets
      if (is.null(self$sheet_names)) {
        return(NA_integer_)
      }

      # input is number
      if (is.numeric(sheet)) {
        badsheet <- !sheet %in% seq_along(self$sheet_names)
        if (any(badsheet)) sheet[badsheet] <- NA_integer_
        return(sheet)
      }

      if (!sheet %in% replaceXMLEntities(self$sheet_names)) {
        return(NA_integer_)
      }

      which(replaceXMLEntities(self$sheet_names) == sheet)
    },

    #' @description
    #' Add worksheet to the `wbWorkbook` object
    #' @param sheet sheet
    #' @param gridLines gridLines
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
      visible <- tolower(as.character(visible))
      visible <- match.arg(visible)
      orientation <- match.arg(orientation, c("portrait", "landscape"))

      # set up so that a single error can be reported on fail
      fail <- FALSE
      msg <- NULL

      private$validate_new_sheet(sheet)
      sheet <- as.character(sheet)

      if (tolower(sheet) %in% tolower(self$sheet_names)) {
        fail <- TRUE
        msg <- c(
          msg,
          sprintf("A worksheet by the name \"%s\" already exists.", sheet),
          "Sheet names must be unique case-insensitive."
        )
      }

      if (!is.logical(gridLines) | length(gridLines) > 1) {
        fail <- TRUE
        msg <- c(msg, "gridLines must be a logical of length 1.")
      }

      if (nchar(sheet) > 31) {
        fail <- TRUE
        msg <- c(
          msg,
          sprintf("sheet \"sheet\" too long.", sheet),
          "Max length is 31 characters."
        )
      }

      if (!is.null(tabColour)) {
        tabColour <- validateColour(tabColour, "Invalid tabColour in add_worksheet.")
      }

      if (!is.numeric(zoom)) {
        fail <- TRUE
        msg <- c(msg, "zoom must be numeric")
      }

      if (!is.null(oddHeader) & length(oddHeader) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("header"))
      }

      if (!is.null(oddFooter) & length(oddFooter) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("footer"))
      }

      if (!is.null(evenHeader) & length(evenHeader) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("evenHeader"))
      }

      if (!is.null(evenFooter) & length(evenFooter) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("evenFooter"))
      }

      if (!is.null(firstHeader) & length(firstHeader) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("firstHeader"))
      }

      if (!is.null(firstFooter) & length(firstFooter) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("firstFooter"))
      }

      vdpi <- as.integer(vdpi)
      hdpi <- as.integer(hdpi)

      if (is.na(vdpi)) {
        fail <- TRUE
        msg <- c(msg, "vdpi must be numeric")
      }

      if (is.na(hdpi)) {
        fail <- TRUE
        msg <- c(msg, "hdpi must be numeric")
      }

      ## Invalid XML characters
      sheet <- replaceIllegalCharacters(sheet)

      if (!missing(sheet)) {
        if (grepl(":", sheet)) {
          fail <- TRUE
          msg <- c(msg, "colon not allowed in sheet names in Excel")
        }
      }

      if (fail) {
        stop(msg, call. = FALSE)
      }

      newSheetIndex <- length(self$worksheets) + 1L
      sheetId <- max_sheet_id(self) # checks for self$worksheet length

      # check for errors ----

      visible <- switch(
        visible,
        true = "visible",
        false = "hidden",
        veryhidden = "veryHidden",
        visible
      )

      # Order matters: if a sheet is added to a blank workbook, we add a default style. If we already have
      # sheets in the workbook, we do not add a new style. This could confuse Excel which will complain.
      # This fixes output of the example in wb_load.
      if (length(self$sheet_names) == 0) {
        # TODO this should live wherever the other default values for an empty worksheet are initialized
        empty_cellXfs <- data.frame(numFmtId = "0", fontId = "0", fillId = "0", borderId = "0", xfId = "0", stringsAsFactors = FALSE)
        self$styles_mgr$styles$cellXfs <- write_xf(empty_cellXfs)
      }

      ##  Add sheet to workbook.xml
      self$append_sheets(
        sprintf(
          '<sheet name="%s" sheetId="%s" state="%s" r:id="rId%s"/>',
          sheet,
          sheetId,
          visible,
          newSheetIndex
        )
      )

      ## append to worksheets list
      self$append("worksheets",
        wbWorksheet$new(
          gridLines   = gridLines,
          tabSelected = newSheetIndex == 1,
          tabColour   = tabColour,
          zoom        = zoom,
          oddHeader   = oddHeader,
          oddFooter   = oddFooter,
          evenHeader  = evenHeader,
          evenFooter  = evenFooter,
          firstHeader = firstHeader,
          firstFooter = firstFooter,
          paperSize   = paperSize,
          orientation = orientation,
          hdpi        = hdpi,
          vdpi        = vdpi
        )
      )


      ## update content_tyes
      ## add a drawing.xml for the worksheet
      if (hasDrawing) {
        self$append("Content_Types", c(
          sprintf('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', newSheetIndex),
          sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', newSheetIndex)
        ))
      } else {
        self$append("Content_Types",
          sprintf(
            '<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            newSheetIndex
          )
        )
      }

      ## Update xl/rels
      self$append("workbook.xml.rels",
        sprintf(
          '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>',
          newSheetIndex
        )
      )

      ## create sheet.rels to simplify id assignment
      self$worksheets_rels[[newSheetIndex]]  <- genBaseSheetRels(newSheetIndex)
      self$drawings_rels[[newSheetIndex]]    <- list()
      self$drawings[[newSheetIndex]]         <- list()
      self$vml_rels[[newSheetIndex]]         <- list()
      self$vml[[newSheetIndex]]              <- list()
      self$isChartSheet[[newSheetIndex]]     <- FALSE
      self$comments[[newSheetIndex]]         <- list()
      self$threadComments[[newSheetIndex]]   <- list()
      self$rowHeights[[newSheetIndex]]       <- list()

      self$append("sheetOrder", as.integer(newSheetIndex))
      self$append("sheet_names", sheet)

      invisible(self)
    },

    # TODO should this be as simple as: wb$wb_add_worksheet(wb$worksheets[[1]]$clone()) ?

    #' @description
    #' Clone a workbooksheet
    #' @param old name of worksheet to clone
    #' @param new name of new worksheet to add
    clone_worksheet = function(old, new) {
      old <- private$get_sheet(old)

      if (tolower(new) %in% tolower(self$sheet_names)) {
        stop("A worksheet by that name already exists! Sheet names must be unique case-insensitive.")
      }

      if (nchar(new) > 31) {
        stop("sheet too long! Max length is 31 characters.")
      }

      if (!is.character(new)) {
        new <- as.character(new)
      }

      ## Invalid XML characters
      new <- replaceIllegalCharacters(new)
      if (grepl(pattern = ":", x = new)) {
        stop("colon not allowed in sheet names in Excel")
      }

      newSheetIndex <- length(self$worksheets) + 1L
      sheetId <- max_sheet_id(self) # checks for length of worksheets


      ## copy visibility from cloned sheet!
      visible <- reg_match0(self$workbook$sheets[[old]], '(?<=state=")[^"]+')

      ##  Add sheet to workbook.xml
      self$append_sheets(
        sprintf(
          '<sheet name="%s" sheetId="%s" state="%s" r:id="rId%s"/>',
          new,
          sheetId,
          visible,
          newSheetIndex
        )
      )

      ## append to worksheets list
      self$append("worksheets", self$worksheets[[old]]$clone())

      ## update content_tyes
      ## add a drawing.xml for the worksheet
      self$append("Content_Types", c(
        sprintf('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', newSheetIndex),
        sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', newSheetIndex)
      ))

      ## Update xl/rels
      self$append("workbook.xml.rels",
        sprintf('<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>', newSheetIndex)
      )

      ## create sheet.rels to simplify id assignment
      self$worksheets_rels[[newSheetIndex]] <- genBaseSheetRels(newSheetIndex)
      self$drawings_rels[[newSheetIndex]] <- self$drawings_rels[[old]]

      # give each chart its own filename (images can re-use the same file, but charts can't)
      self$drawings_rels[[newSheetIndex]] <-
        # TODO Can this be simplified?  There's a bit going on here
        vapply(
          self$drawings_rels[[newSheetIndex]],
          function(rl) {
            # is rl here a length of 1?
            stopifnot(length(rl) == 1L) # lets find out...  if this fails, just remove it
            chartfiles <- reg_match(rl, "(?<=charts/)chart[0-9]+\\.xml")

            for (cf in chartfiles) {
              chartid <- length(.self$charts) + 1L
              newname <- stri_join("chart", chartid, ".xml")
              fl <- self$charts[cf]

              # Read the chartfile and adjust all formulas to point to the new
              # sheet name instead of the clone source
              # The result is saved to a new chart xml file
              newfl <- file.path(dirname(fl), newname)

              self$charts[newname] <- newfl

              chart <- read_xml(fl, pointer = FALSE)

              chart <- gsub(
                stri_join("(?<=')", self$sheet_names[[old]], "(?='!)"),
                stri_join("'", new, "'"),
                chart,
                perl = TRUE
              )

              chart <- gsub(
                stri_join("(?<=[^A-Za-z0-9])", .self$sheet_names[[old]], "(?=!)"),
                stri_join("'", new, "'"),
                chart,
                perl = TRUE
              )

              writeLines(chart, newfl)

              self$append("Content_Types",
                sprintf('<Override PartName="/xl/charts/%s" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', newname)
              )

              rl <- gsub(stri_join("(?<=charts/)", cf), newname, rl, perl = TRUE)
            }
            rl
          },
          NA_character_,
          USE.NAMES = FALSE
        )
      # The IDs in the drawings array are sheet-specific, so within the new
      # cloned sheet the same IDs can be used => no need to modify drawings
      self$drawings[[newSheetIndex]]       <- self$drawings[[old]]
      self$vml_rels[[newSheetIndex]]       <- self$vml_rels[[old]]
      self$vml[[newSheetIndex]]            <- self$vml[[old]]
      self$isChartSheet[[newSheetIndex]]   <- self$isChartSheet[[old]]
      self$comments[[newSheetIndex]]       <- self$comments[[old]]
      self$threadComments[[newSheetIndex]] <- self$threadComments[[old]]
      self$rowHeights[[newSheetIndex]]     <- self$rowHeights[[old]]

      self$append("sheetOrder", as.integer(newSheetIndex))
      self$append("sheet_names", new)


      ############################
      ## TABLES
      ## ... are stored in the $tables list, with the name and sheet as attr
      ## and in the worksheets[]$tableParts list. We also need to adjust the
      ## worksheets_rels and set the content type for the new table

      tbls <- self$tables[attr(self$tables, "sheet") == old]

      for (t in tbls) {
        # Extract table name, displayName and ID from the xml
        oldname     <- reg_match0(t, '(?<= name=")[^"]+')
        olddispname <- reg_match0(t, '(?<= displayName=")[^"]+')
        oldid       <- reg_match0(t, '(?<= id=")[^"]+')
        ref         <- reg_match0(t, '(?<= ref=")[^"]+')

        # Find new, unused table names by appending _n, where n=1,2,...
        n <- 0
        while (stri_join(oldname, "_", n) %in% attr(self$tables, "tableName")) {
          n <- n + 1
        }

        newname <- stri_join(oldname, "_", n)
        newdispname <- stri_join(olddispname, "_", n)
        newid <- as.character(length(self$tables) + 3L)

        # Use the table definition from the cloned sheet and simply replace the names
        newt <- t
        newt <- gsub(
          stri_join(' name="', oldname, '"'),
          stri_join(' name="', newname, '"'),
          newt
        )
        newt <- gsub(
          stri_join(' displayName="', olddispname, '"'),
          stri_join(' displayName="', newdispname, '"'),
          newt
        )
        newt <- gsub(
          stri_join('(<table [^<]* id=")', oldid, '"'),
          stri_join("\\1", newid, '"'),
          newt
        )

        oldtables <- self$tables
        self$append("tables", newt)
        names(self$tables)             <- c(names(oldtables), ref)
        attr(self$tables, "sheet")     <- c(attr(oldtables, "sheet"), newSheetIndex)
        attr(self$tables, "tableName") <- c(attr(oldtables, "tableName"), newname)

        oldparts                                                              <- self$worksheets[[newSheetIndex]]$tableParts
        self$worksheets[[newSheetIndex]]$tableParts                           <- c(oldparts, sprintf('<tablePart r:id="rId%s"/>', newid))
        attr(self$worksheets[[newSheetIndex]]$tableParts, "tableName")        <- c(attr(oldparts, "tableName"), newname)
        names(attr(self$worksheets[[newSheetIndex]]$tableParts, "tableName")) <- c(names(attr(oldparts, "tableName")), ref)

        self$append("Content_Types", sprintf('
          <Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
          newid
        ))

        self$append("tables.xml.rels", "")

        private$append_sheet_rels(newSheetIndex, sprintf(
          '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',
          newid,
          newid
        ))
      }

      # TODO: The following items are currently NOT copied/duplicated for the cloned sheet:
      #   - Comments
      #   - Pivot tables

      # invisible(newSheetIndex)
      invisible(self)
    },

    #' @description
    #' Add a chart sheet to the workbook
    #' @param sheet sheet
    #' @param tabColour tabColour
    #' @param zoom zoom
    #' @return The `wbWorkbook` object, invisibly
    addChartSheet = function(sheet, tabColour = NULL, zoom = 100) {
      # TODO private$new_sheet_index()?
      newSheetIndex <- length(self$worksheets) + 1L
      sheetId <- max_sheet_id(self) # checks for length of worksheets

      ##  Add sheet to workbook.xml
      self$append_sheets(
        sprintf(
          '<sheet name="%s" sheetId="%s" r:id="rId%s"/>',
          sheet,
          sheetId,
          newSheetIndex
        )
      )

      ## append to worksheets list
      self$append("worksheets",
        wbChartSheet$new(
          tabSelected = newSheetIndex == 1,
          tabColour = tabColour,
          zoom = zoom
        )
      )

      self$append("sheet_names", sheet)

      ## update content_tyes
      self$append("Content_Types",
        sprintf(
          '<Override PartName="/xl/chartsheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>',
          newSheetIndex
        )
      )

      ## Update xl/rels
      self$append("workbook.xml.rels",
        sprintf(
          '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet%s.xml"/>',
          newSheetIndex
        )
      )

      ## add a drawing.xml for the worksheet
      self$append("Content_Types",
        sprintf(
          '<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
          newSheetIndex
        )
      )

      ## create sheet.rels to simplify id assignment
      self$worksheets_rels[[newSheetIndex]]  <- genBaseSheetRels(newSheetIndex)
      self$drawings_rels[[newSheetIndex]]    <- list()
      self$drawings[[newSheetIndex]]         <- list()
      self$isChartSheet[[newSheetIndex]]     <- TRUE
      self$rowHeights[[newSheetIndex]]       <- list()
      self$vml_rels[[newSheetIndex]]         <- list()
      self$vml[[newSheetIndex]]              <- list()
      self$append("sheetOrder", newSheetIndex)

      # invisible(newSheetIndex)
      invisible(self)
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
      assert_class(path, "character")
      assert_class(overwrite, "logical")

      if (file.exists(path) & !overwrite) {
        stop("File already exists!")
      }

      ## temp directory to save XML files prior to compressing
      tmpDir <- file.path(tempfile(pattern = "workbookTemp_"))
      on.exit(unlink(tmpDir, recursive = TRUE), add = TRUE)

      if (file.exists(tmpDir)) {
        unlink(tmpDir, recursive = TRUE, force = TRUE)
      }

      success <- dir.create(path = tmpDir, recursive = TRUE)
      if (!success) {
        stop(sprintf("Failed to create temporary directory '%s'", tmpDir))
      }

      private$preSaveCleanUp()

      nSheets         <- length(self$worksheets)
      nThemes         <- length(self$theme)
      nPivots         <- length(self$pivotDefinitions)
      nSlicers        <- length(self$slicers)
      nComments       <- sum(lengths(self$comments) > 0)
      nThreadComments <- sum(lengths(self$threadComments) > 0)
      nPersons        <- length(self$persons)
      nVML            <- sum(lengths(self$vml) > 0)

      relsDir         <- dir_create(tmpDir, "_rels")
      docPropsDir     <- dir_create(tmpDir, "docProps")
      xlDir           <- dir_create(tmpDir, "xl")
      xlrelsDir       <- dir_create(tmpDir, "xl", "_rels")
      xlTablesDir     <- dir_create(tmpDir, "xl", "tables")
      xlTablesRelsDir <- dir_create(xlTablesDir, "_rels")

      if (length(self$media)) {
        xlmediaDir <- dir_create(tmpDir, "xl", "media")
      }

      ## will always have a theme
      xlthemeDir <- dir_create(tmpDir, "xl", "theme")

      if (is.null(self$theme)) {
        con <- file(file.path(xlthemeDir, "theme1.xml"), open = "wb")
        writeBin(charToRaw(genBaseTheme()), con)
        close(con)
      } else {
        # TODO replace with seq_len() or seq_along()
        lapply(seq_len(nThemes), function(i) {
          con <- file(file.path(xlthemeDir, stri_join("theme", i, ".xml")), open = "wb")
          writeBin(charToRaw(pxml(self$theme[[i]])), con)
          close(con)
        })
      }


      ## will always have drawings
      xlworksheetsDir     <- dir_create(tmpDir, "xl", "worksheets")
      xlworksheetsRelsDir <- dir_create(tmpDir, "xl", "worksheets", "_rels")
      xldrawingsDir       <- dir_create(tmpDir, "xl", "drawings")
      xldrawingsRelsDir   <- dir_create(tmpDir, "xl", "drawings", "_rels")

      ## charts
      if (length(self$charts)) {
        file.copy(
          from = dirname(self$charts[1]),
          to = file.path(tmpDir, "xl"),
          recursive = TRUE
        )
      }


      ## xl/comments.xml
      if (nComments > 0 | nVML > 0) {
        # TODO use seq_len() or seq_along()?
        for (i in seq_len(nSheets)) {
          if (length(self$comments[[i]])) {
            fn <- sprintf("comments%s.xml", i)

            write_comment_xml(
              comment_list = self$comments[[i]],
              file_name = file.path(tmpDir, "xl", fn)
            )
          }
        }

        private$writeDrawingVML(xldrawingsDir)
      }

      ## Threaded Comments xl/threadedComments/threadedComment.xml
      if (nThreadComments > 0) {
        xlThreadComments <- dir_create(tmpDir, "xl", "threadedComments")

        for (i in seq_len(nSheets)) {
          if (length(self$threadComments[[i]])) {
            fl <- self$threadComments[[i]]
            file.copy(
              from = fl,
              to = file.path(xlThreadComments, basename(fl)),
              overwrite = TRUE,
              copy.date = TRUE
            )
          }
        }
      }

      ## xl/persons/person.xml
      if (nPersons) {
        personDir <- dir_create(tmpDir, "xl", "persons")
        file.copy(self$persons, personDir, overwrite = TRUE)
      }



      if (length(self$embeddings)) {
        embeddingsDir <- dir_create(tmpDir, "xl", "embeddings")
        for (fl in self$embeddings) {
          file.copy(fl, embeddingsDir, overwrite = TRUE)
        }
      }


      if (nPivots > 0) {
        # TODO consider just making a function to create a bunch of directories
        # and return as a named list?  Easier/cleaner than checking for each
        # element if we just go seq_along()?
        pivotTablesDir     <- dir_create(tmpDir, "xl", "pivotTables")
        pivotTablesRelsDir <- dir_create(tmpDir, "xl", "pivotTables", "_rels")
        pivotCacheDir      <- dir_create(tmpDir, "xl", "pivotCache")
        pivotCacheRelsDir  <- dir_create(tmpDir, "xl", "pivotCache", "_rels")

        file_copy_wb_save(self$pivotTables,          "pivotTable%i.xml",                pivotTablesDir)
        file_copy_wb_save(self$pivotDefinitions,     "pivotCacheDefinition%i.xml",      pivotCacheDir)
        file_copy_wb_save(self$pivotRecords,         "pivotCacheRecords%i.xml",         pivotCacheDir)
        file_copy_wb_save(self$pivotDefinitionsRels, "pivotCacheDefinition%i.xml.rels", pivotCacheRelsDir)

        for (i in seq_along(self$pivotTables.xml.rels)) {
          write_file(
            body = self$pivotTables.xml.rels[[i]],
            fl = file.path(pivotTablesRelsDir, sprintf("pivotTable%s.xml.rels", i))
          )
        }
      }

      ## slicers
      if (nSlicers) {
        slicersDir      <- dir_create(tmpDir, "xl", "slicers")
        slicerCachesDir <- dir_create(tmpDir, "xl", "slicerCaches")

        # for (i in which(nchar(self$slicers > 1))) {
        for (i in which(nzchar(self$slicers))) {
          file.copy(self$slicers[i], file.path(slicersDir, sprintf("slicer%s.xml", i)), overwrite = TRUE, copy.date = TRUE)
        }

        for (i in seq_along(self$slicerCaches)) {
          write_file(
            body = self$slicerCaches[[i]],
            fl = file.path(slicerCachesDir, sprintf("slicerCache%s.xml", i))
          )
        }
      }


      ## Write content

      ## write .rels
      write_file(
        head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n',
        body = '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
        tail = "</Relationships>",
        fl = file.path(relsDir, ".rels")
      )

      app <- "<Application>Microsoft Excel</Application>"
      # further protect argument (might be extended with: <ScaleCrop>, <HeadingPairs>, <TitlesOfParts>, <LinksUpToDate>, <SharedDoc>, <HyperlinksChanged>, <AppVersion>)
      if (!is.null(self$apps)) app <- paste0(app, self$apps)

      ## write app.xml
      if (length(self$app) == 0) {
        write_file(
          head = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
          body = app,
          tail = "</Properties>",
          fl = file.path(docPropsDir, "app.xml")
        )
      } else {
        write_file(
          head = '',
          body = pxml(self$app),
          tail = '',
          fl = file.path(docPropsDir, "app.xml")
        )
      }

      ## write core.xml
      write_file(
        head = "",
        body = pxml(self$core),
        tail = "",
        fl = file.path(docPropsDir, "core.xml")
      )

      ## write workbook.xml.rels
      write_file(
        head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        body = pxml(self$workbook.xml.rels),
        tail = "</Relationships>",
        fl = file.path(xlrelsDir, "workbook.xml.rels")
      )

      ## write tables

      ## Content types has entries of all xml files in the workbook
      ct <- self$Content_Types
      ## update tables in content types (some have been added, some removed, get the final state)
      default <- xml_node(ct, "Default")
      override <- rbindlist(xml_attr(ct, "Override"))
      override$typ <- gsub(".xml$", "", basename(override$PartName))
      override <- override[!grepl("table", override$typ), ]
      override$typ <- NULL

      # TODO remove length() check since we have seq_along()
      if (length(unlist(self$tables, use.names = FALSE))) {

        # TODO get table Id from table entry
        table_ids <- function() {
          relship <- rbindlist(xml_attr(unlist(self$worksheets_rels), "Relationship"))
          relship$typ <- basename(relship$Type)
          relship$tid <- as.numeric(gsub("\\D+", "", relship$Target))
          sort(relship$tid[relship$typ == "table"])
        }

        tab_ids <- table_ids()
        for (i in seq_along(tab_ids)) {

          idx <- attr(self$tables, "sheet") > 0

          if (!grepl("openxlsx_deleted", attr(self$tables, "tableName")[idx][i], fixed = TRUE)) {
            write_file(
              body = pxml(unlist(self$tables[idx], use.names = FALSE)[[i]]),
              fl = file.path(xlTablesDir, sprintf("table%s.xml", tab_ids[[i]]))
            )

            ## add entry to content_types as well
            override <- rbind(
              override,
              # new entry for table
              c("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                sprintf("/xl/tables/table%s.xml", tab_ids[[i]]))
            )

            if (self$tables.xml.rels[[i]] != "") {
              write_file(
                body = self$tables.xml.rels[[i]],
                fl = file.path(xlTablesRelsDir, sprintf("table%s.xml.rels", tab_ids[[i]]))
              )
            }
          }
        }
      }

      ## ct is updated as xml
      ct <- c(default, df_to_xml(name = "Override", df_col = override[c("PartName", "ContentType")]))


      ## write query tables
      if (length(self$queryTables)) {
        xlqueryTablesDir <- dir_create(tmpDir, "xl", "queryTables")

        for (i in seq_along(self$queryTables)) {
          write_file(
            body = self$queryTables[[i]],
            fl = file.path(xlqueryTablesDir, sprintf("queryTable%s.xml", i))
          )
        }
      }

      ## connections
      if (length(self$connections)) {
        write_file(body = self$connections, fl = file.path(xlDir, "connections.xml"))
      }

      ## externalLinks
      if (length(self$externalLinks)) {
        externalLinksDir <- dir_create(tmpDir, "xl", "externalLinks")

        for (i in seq_along(self$externalLinks)) {
          write_file(
            body = self$externalLinks[[i]],
            fl = file.path(externalLinksDir, sprintf("externalLink%s.xml", i))
          )
        }
      }

      ## externalLinks rels
      if (length(self$externalLinksRels)) {
        externalLinksRelsDir <- dir_create(tmpDir, "xl", "externalLinks", "_rels")

        for (i in seq_along(self$externalLinksRels)) {
          write_file(
            body = self$externalLinksRels[[i]],
            fl = file.path(
              externalLinksRelsDir,
              sprintf("externalLink%s.xml.rels", i)
            )
          )
        }
      }

      ## media (copy file from origin to destination)
      # TODO replace with seq_along()
      for (x in self$media) {
        file.copy(x, file.path(xlmediaDir, names(self$media)[which(self$media == x)]))
      }

      ## VBA Macro
      if (!is.null(self$vbaProject)) {
        file.copy(self$vbaProject, xlDir)
      }

      ## write worksheet, worksheet_rels, drawings, drawing_rels
      private$writeSheetDataXML(
        xldrawingsDir,
        xldrawingsRelsDir,
        xlworksheetsDir,
        xlworksheetsRelsDir
      )

      ## write sharedStrings.xml
      if (length(self$sharedStrings)) {
        write_file(
          head = sprintf(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%s" uniqueCount="%s">',
            length(self$sharedStrings),
            attr(self$sharedStrings, "uniqueCount")
          ),
          #body = stri_join(set_sst(attr(sharedStrings, "text")), collapse = "", sep = " "),
          body = stri_join(self$sharedStrings, collapse = "", sep = ""),
          tail = "</sst>",
          fl = file.path(xlDir, "sharedStrings.xml")
        )
      } else {
        ## Remove relationship to sharedStrings
        ct <- ct[!grepl("sharedStrings", ct)]
      }

      for (draw in seq_along(self$drawings)) {
          if (length(self$drawings[[draw]])) {
              ct <- c(ct,
                sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', draw)
              )
          }
      }

      if (nComments > 0) {
        ct <- c(
          ct,
          # TODO this default extension is most likely wrong here and should be set when searching for and writing the vml entrys
          '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
          sprintf('<Override PartName="/xl/comments%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>', seq_len(nComments)
          )
        )
      }

      self$Content_Types <- ct

      ## write [Content_type]
      write_file(
        head = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        body = pxml(unique(ct)),
        tail = "</Types>",
        fl = file.path(tmpDir, "[Content_Types].xml")
      )


      styleXML <- self$styles_mgr$styles
      if (length(styleXML$numFmts)) {
        styleXML$numFmts <-
          stri_join(
            sprintf('<numFmts count="%s">', length(styleXML$numFmts)),
            pxml(styleXML$numFmts),
            "</numFmts>"
          )
      }
      styleXML$fonts <-
        stri_join(
          sprintf('<fonts count="%s">', length(styleXML$fonts)),
          pxml(styleXML$fonts),
          "</fonts>"
        )
      styleXML$fills <-
        stri_join(
          sprintf('<fills count="%s">', length(styleXML$fills)),
          pxml(styleXML$fills),
          "</fills>"
        )
      styleXML$borders <-
        stri_join(
          sprintf('<borders count="%s">', length(styleXML$borders)),
          pxml(styleXML$borders),
          "</borders>"
        )
      styleXML$cellStyleXfs <-
        c(
          sprintf('<cellStyleXfs count="%s">', length(styleXML$cellStyleXfs)),
          pxml(styleXML$cellStyleXfs),
          "</cellStyleXfs>"
        )
      styleXML$cellXfs <-
        stri_join(
          sprintf('<cellXfs count="%s">', length(styleXML$cellXfs)),
          paste0(styleXML$cellXfs, collapse = ""),
          "</cellXfs>"
        )
      styleXML$cellStyles <-
        stri_join(
          sprintf('<cellStyles count="%s">', length(styleXML$cellStyles)),
          pxml(styleXML$cellStyles),
          "</cellStyles>"
        )
      # styleXML$cellStyles <-
      #   stri_join(
      #     pxml(self$styles_mgr$tableStyles)
      #   )
      # TODO
      # tableStyles
      # extLst

      styleXML$dxfs <-
        if (length(styleXML$dxfs)) {
          stri_join(
            sprintf('<dxfs count="%s">', length(styleXML$dxfs)),
            stri_join(unlist(styleXML$dxfs), sep = " ", collapse = ""),
            "</dxfs>"
          )
        } else {
          '<dxfs count="0"/>'
        }

      ## write styles.xml
      #if (class(self$styles_xml) == "uninitializedField") {
      write_file(
        head = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">',
        body = pxml(styleXML),
        tail = "</styleSheet>",
        fl = file.path(xlDir, "styles.xml")
      )
      #} else {
      #  write_file(
      #    head = '',
      #    body = self$styles_xml,
      #    tail = '',
      #    fl = file.path(xlDir, "styles.xml")
      #  )
      #}

      if (length(self$calcChain)) {
        write_file(
          head = '',
          body = pxml(self$calcChain),
          tail = "",
          fl = file.path(xlDir, "calcChain.xml")
        )
      }

      ## write workbook.xml
      workbookXML <- self$workbook
      workbookXML$sheets <- stri_join("<sheets>", pxml(workbookXML$sheets), "</sheets>")

      if (length(workbookXML$definedNames)) {
        workbookXML$definedNames <- stri_join("<definedNames>", pxml(workbookXML$definedNames), "</definedNames>" )
      }

      write_file(
        head = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">',
        body = pxml(workbookXML),
        tail = "</workbook>",
        fl = file.path(xlDir, "workbook.xml")
      )

      ## Need to reset sheet order to allow multiple savings
      self$workbook$sheets <- self$workbook$sheets[order(self$sheetOrder)]

      ## compress to xlsx

      # TODO make self$vbaProject be TRUE/FALSE
      tmpFile <- tempfile(tmpdir = tmpDir, fileext = if (isTRUE(self$vbaProject)) ".xlsm" else ".xlsx")

      ## zip it
      zip::zip(
        zipfile = tmpFile,
        files = list.files(tmpDir, full.names = FALSE),
        recurse = TRUE,
        compression_level = getOption("openxlsx2.compresssionevel", 6),
        include_directories = FALSE,
        # change the working directory for this
        root = tmpDir,
        # change default to match historical zipr
        mode = "cherry-pick"
      )

      # Copy file; stop if failed
      if (!file.copy(from = tmpFile, to = path, overwrite = overwrite, copy.mode = FALSE)) {
        stop("Failed to save workbook")
      }

      # (re)assign file path (if successful)
      self$path <- path
      invisible(self)
    },

    #' @description open wbWorkbook in Excel.
    #' @details minor helper wrapping xl_open which does the entire same thing
    #' @return The `wbWorksheetObject`, invisibly
    open = function() {
      tmp <- temp_xlsx()
      self$save(tmp)
      xl_open(tmp)
      invisible(self)
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
      sheet,
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

      ## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
      last_table_id <- function() {
        z <- 0
        relship <- rbindlist(xml_attr(unlist(self$worksheets_rels), "Relationship"))
        relship$typ <- basename(relship$Type)
        relship$tid <- as.numeric(gsub("\\D+", "", relship$Target))
        if (any(relship$typ == "table"))
          z <- max(relship$tid[relship$typ == "table"])

        z
      }

      id <- as.character(last_table_id() + 1) # otherwise will start at 0 for table 1 length indicates the last known
      sheet <- private$get_sheet(sheet)
      rid <- length(xml_node(self$worksheets_rels[[sheet]], "Relationship")) + 1

      nms <- names(self$tables)
      tSheets <- attr(self$tables, "sheet")
      tNames <- attr(self$tables, "tableName")


      ### autofilter
      autofilter <- if (withFilter) {
        xml_node_create(xml_name = "autoFilter", xml_attributes = c(ref = ref))
      }

      ### tableColumn
      tableColumn <- sapply(colNames, function(x) {
        id <- which(colNames %in% x)
        xml_node_create("tableColumn", xml_attributes = c(id = id, name = x))
      })

      tableColumns <- xml_node_create(
        xml_name = "tableColumns",
        xml_children = tableColumn,
        xml_attributes = c(count = as.character(length(colNames)))
      )


      ### tableStyleInfo
      tablestyle_attr <- c(
        name              = tableStyle,
        showFirstColumn   = as.integer(showFirstColumn),
        showLastColumn    = as.integer(showLastColumn),
        showRowStripes    = as.integer(showRowStripes),
        showColumnStripes = as.integer(showColumnStripes)
      )

      tableStyleXML <- xml_node_create(xml_name = "tableStyleInfo", xml_attributes = tablestyle_attr)


      ### full table
      table_attrs <- c(
        xmlns          = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        `xmlns:mc`     = "http://schemas.openxmlformats.org/markup-compatibility/2006",
        id             = id,
        name           = tableName,
        displayName    = tableName,
        ref            = ref,
        totalsRowCount = totalsRowCount,
        totalsRowShown = "0"
        #headerRowDxfId="1"
      )

      self$append("tables",
        xml_node_create(
          xml_name = "table",
          xml_children = c(autofilter, tableColumns, tableStyleXML),
          xml_attributes = table_attrs
        )
      )

      names(self$tables) <- c(nms, ref)
      attr(self$tables, "sheet") <- c(tSheets, sheet)
      attr(self$tables, "tableName") <- c(tNames, tableName)

      private$append_sheet_field(sheet, "tableParts", sprintf('<tablePart r:id="rId%s"/>', rid))
      ok <- tSheets == sheet & !grepl("openxlsx_deleted", tNames)
      attr(self$worksheets[[sheet]]$tableParts, "tableName") <- c(tNames[ok], tableName)

      ## create a table.xml.rels
      self$append("tables.xml.rels", "")

      ## update worksheets_rels
      # TODO do we have to worry about exisiting table relationships?
      private$append_sheet_rels(sheet, sprintf(
        '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',
        rid,
        id
      ))
      invisible(self)
    },


    #' @description
    #' Create a font node from a style
    #' @param style style
    #' @return The font node as xml?
    createFontNode = function(style) {
      # Nothing is assigned to self, so this is fine to not return self
      # TODO pull out from public methods
      # assert_style(style)
      baseFont <- self$get_base_font()

      # Create font name
      paste_c(
        "<font>",

        ## size
        if (is.null(style$fontSize[[1]])) {
          sprintf('<sz %s="%s"/>', names(baseFont$size), baseFont$size)
        } else {
          sprintf('<sz %s="%s"/>', names(style$fontSize), style$fontSize)
        },

        # colour
        if (is.null(style$fontColour[[1]])) {
          sprintf('<color %s="%s"/>', names(baseFont$colour), baseFont$colour)
        } else {
          new <- sprintf('%s="%s"', names(style$fontColour), style$fontColour)
          sprintf("<color %s/>", stri_join(new, sep = " ", collapse = " "))
        },

        # font name
        if (is.null(style$fontName[[1]])) {
          sprintf('<name %s="%s"/>', names(baseFont$name), baseFont$name)
        } else {
          sprintf('<name %s="%s"/>', names(style$fontName), style$fontName)
        },

        # new font name and return id
        sprintf('<family val = "%s"/>', style$fontFamily),
        sprintf('<scheme val = "%s"/>', style$fontScheme),

        if ("BOLD"       %in% style$fontDecoration) "<b/>",
        if ("ITALIC"     %in% style$fontDecoration) "<i/>",
        if ("UNDERLINE"  %in% style$fontDecoration) '<u val="single"/>',
        if ("UNDERLINE2" %in% style$fontDecoration) '<u val="double"/>',
        if ("STRIKEOUT"  %in% style$fontDecoration) "<strike/>",

        "</font>"
      )
    },

    #' @description
    #' Get the base font
    #' @return A list of of the font
    get_base_font = function() {
      baseFont <- self$styles_mgr$styles$fonts[[1]]

      sz     <- unlist(xml_attr(baseFont, "font", "sz"))
      colour <- unlist(xml_attr(baseFont, "font", "color"))
      name   <- unlist(xml_attr(baseFont, "font", "name"))

      if (length(sz[[1]]) == 0) {
        sz <- list("val" = "10")
      } else {
        sz <- as.list(sz)
      }

      if (length(colour[[1]]) == 0) {
        colour <- list("rgb" = "#000000")
      } else {
        colour <- as.list(colour)
      }

      if (length(name[[1]]) == 0) {
        name <- list("val" = "Calibri")
      } else {
        name <- as.list(name)
      }

      list(
        size   = sz,
        colour = colour,
        name   = name
      )
    },

    #' @description
    #' Get the base font
    #' @param fontSize fontSize
    #' @param fontColour fontColour
    #' @param fontName fontName
    #' @return The `wbWorkbook` object
    set_base_font = function(fontSize = 11, fontColour = "black", fontName = "Calibri") {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      if (fontSize < 0) stop("Invalid fontSize")
      fontColour <- validateColour(fontColour)

      self$styles_mgr$styles$fonts[[1]] <- sprintf(
        '<font><sz val="%s"/><color rgb="%s"/><name val="%s"/></font>',
        fontSize,
        fontColour,
        fontName
      )
    },

    #' @description
    #' Sets a sheet name
    #' @param sheet Old sheet name
    #' @param name New sheet name
    #' @return The `wbWorkbook` object, invisibly
    setSheetName = function(sheet, name) {
      # TODO assert sheet class?
      if (name %in% self$sheet_names) {
        stop(sprintf("Sheet %s already exists!", name))
      }

      sheet <- private$get_sheet(sheet)

      oldName <- self$sheet_names[[sheet]]
      self$sheet_names[[sheet]] <- name

      ## Rename in workbook
      sheetId <- get_sheet_id(self, sheet)
      rId <- get_r_id(self, sheet)
      self$workbook$sheets[[sheet]] <-
        sprintf(
          '<sheet name="%s" sheetId="%s" r:id="rId%s"/>',
          name,
          sheetId,
          rId
        )

      ## rename defined names
      if (length(self$workbook$definedNames)) {
        belongTo <- get_named_regions(self)$sheets
        toChange <- belongTo == oldName
        if (any(toChange)) {
          name <- sprintf("'%s'", name)
          tmp <-
            gsub(oldName, name, self$workbook$definedName[toChange], fixed = TRUE)
          tmp <- gsub("'+", "'", tmp)
          self$workbook$definedNames[toChange] <- tmp
        }
      }

      invisible(self)
    },

    #' @description
    #' Sets a row height for a sheet
    #' @param sheet sheet
    #' @param rows rows
    #' @param heights heights
    #' @return The `wbWorkbook` object, invisibly
    set_row_heights = function(sheet, rows, heights) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      sheet <- private$get_sheet(sheet)
      # TODO move to wbWorksheet method
      # TODO consider reworking rowHeights
      # self$worksheets[[sheet]]$set_row_heights(rows = rows, heights = heights)
      # invisible(self)

      if (length(rows) > length(heights)) {
        heights <- rep(heights, length.out = length(rows))
      }

      if (length(heights) > length(rows)) {
        stop("Greater number of height values than rows.")
      }

      ## Remove duplicates
      ok <- !duplicated(rows)
      heights <- heights[ok]
      rows <- rows[ok]

      heights <- as.character(as.numeric(heights))
      names(heights) <- rows

      ## remove any conflicting heights
      flag <- names(self$rowHeights[[sheet]]) %in% rows
      if (any(flag)) {
        self$rowHeights[[sheet]] <- self$rowHeights[[sheet]][!flag]
      }

      nms <- c(names(self$rowHeights[[sheet]]), rows)
      allRowHeights <- unlist(c(self$rowHeights[[sheet]], heights))
      names(allRowHeights) <- nms

      allRowHeights <-
        allRowHeights[order(as.integer(names(allRowHeights)))]

      self$rowHeights[[sheet]] <- allRowHeights
      invisible(self)
    },

    #' @description
    #' Sets a row height for a sheet
    #' @param sheet sheet
    #' @param rows rows
    #' @return The `wbWorkbook` object, invisibly
    remove_row_heights = function(sheet, rows) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      sheet <- private$get_sheet(sheet)

      customRows <- as.integer(names(self$rowHeights[[sheet]]))
      removeInds <- which(customRows %in% rows)

      if (length(removeInds)) {
        self$rowHeights[[sheet]] <- self$rowHeights[[sheet]][-removeInds]
      }

      self
    },

    ## columns ----

    #' description
    #' creates column object for worksheet
    #' @param sheet sheet
    #' @param n n
    #' @param beg beg
    #' @param end end
    createCols = function(sheet, n, beg, end) {
       sheet <- private$get_sheet(sheet)
       self$worksheets[[sheet]]$cols_attr <- df_to_xml("col", empty_cols_attr(n, beg, end))
    },

    #' @description
    #' Group cols
    #' @param sheet sheet
    #' @param cols cols
    #' @param collapsed collapsed
    #' @param levels levels
    #' @return The `wbWorkbook` object, invisibly
    group_cols = function(sheet, cols, collapsed = FALSE, levels = NULL) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      sheet <- private$get_sheet(sheet)

      if (length(collapsed) > length(cols)) {
        stop("Collapses argument is of greater length than number of cols.")
      }

      if (!is.logical(collapsed)) {
        stop("Collapses should be a logical value (TRUE/FALSE).")
      }

      if (any(cols < 1L)) {
        stop("Invalid rows entered (<= 0).")
      }

      collapsed <- rep(as.character(as.integer(collapsed)), length.out = length(cols))
      levels <- levels %||% rep("1", length(cols))

      # Remove duplicates
      ok <- !duplicated(cols)
      collapsed <- collapsed[ok]
      levels    <- levels[ok]
      cols      <- cols[ok]

      # fetch the row_attr data.frame
      col_attr <- self$worksheets[[sheet]]$unfold_cols()

      if (NROW(col_attr) == 0) {
        # TODO should this be a warning?  Or an error?
        message("worksheet has no columns. please create some with createCols")
      }

      # reverse to make it easier to get the fist
      cols_rev <- rev(cols)

      # get the selection based on the col_attr frame.

      # the first n -1 cols get outlineLevel
      select <- col_attr$min %in% as.character(cols_rev[-1])
      if (length(select)) {
        col_attr$outlineLevel[select] <- as.character(levels[-1])
        col_attr$collapsed[select] <- as.character(as.integer(collapsed[-1]))
        col_attr$hidden[select] <- as.character(as.integer(collapsed[-1]))
      }

      # the n-th row gets only collapsed
      select <- col_attr$min %in% as.character(cols_rev[1])
      if (length(select)) {
        col_attr$collapsed[select] <- as.character(as.integer(collapsed[1]))
      }

      self$worksheets[[sheet]]$fold_cols(col_attr)


      # check if there are valid outlineLevel in col_attr and assign outlineLevelRow the max outlineLevel (thats in the documentation)
      if (any(col_attr$outlineLevel != "")) {
        self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelCol = as.character(max(as.integer(col_attr$outlineLevel), na.rm = TRUE))))
      }

      invisible(self)
    },

    #' @description ungroup cols
    #' @param sheet sheet
    #' @param cols = cols
    #' @returns The `wbWorkbook` object
    ungroup_cols = function(sheet, cols) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      sheet <- private$get_sheet(sheet)

      # check if any rows are selected
      if (any(cols < 1L)) {
        stop("Invalid cols entered (<= 0).")
      }

      # fetch the cols_attr data.frame
      col_attr <- self$worksheets[[sheet]]$unfold_cols()

      # get the selection based on the col_attr frame.
      select <- col_attr$min %in% as.character(cols)

      if (length(select)) {
        col_attr$outlineLevel[select] <- ""
        col_attr$collapsed[select] <- ""
        # TODO only if unhide = TRUE
        col_attr$hidden[select] <- ""
        self$worksheets[[sheet]]$fold_cols(col_attr)
      }

      # If all outlineLevels are missing: remove the outlineLevelCol attribute. Assigning "" will remove the attribute
      if (all(col_attr$outlineLevel == "")) {
        self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelCol = ""))
      } else {
        self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelCol = as.character(max(as.integer(col_attr$outlineLevel)))))
      }

      self
    },

    #' @description Remove row heights from a worksheet
    #' @param sheet A name or index of a worksheet
    #' @param cols Indices of columns to remove custom width (if any) from.
    #' @return The `wbWorkbook` object, invisibly
    remove_col_widths = function(sheet, cols) {
      sheet <- private$get_sheet(sheet)
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      if (!is.numeric(cols)) {
        cols <- col2int(cols)
      }

      customCols <- as.integer(names(self$colWidths[[sheet]]))
      removeInds <- which(customCols %in% cols)
      if (length(removeInds)) {
        remainingCols <- customCols[-removeInds]
        if (length(remainingCols) == 0) {
          self$colWidths[[sheet]] <- list()
        } else {
          rem_widths <- self$colWidths[[sheet]][-removeInds]
          names(rem_widths) <- as.character(remainingCols)
          self$colWidths[[sheet]] <- rem_widths
        }
      }

      invisible(self)
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
    set_col_widths = function(sheet, cols, widths = 8.43, hidden = FALSE) {
      sheet <- private$get_sheet(sheet)

      # should do nothing if the cols' length is zero
      # TODO why would cols ever be 0?  Can we just signal this as an error?
      if (length(cols) == 0L) {
        return(self)
      }

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
      ok <- !duplicated(cols)
      widths <- widths[ok]
      hidden <- hidden[ok]
      cols <- cols[ok]

      col_df <- self$worksheets[[sheet]]$unfold_cols()

      if (any(widths == "auto")) {

        df <- wb_to_df(self, sheet = sheet, cols = cols, colNames = FALSE)
        # TODO format(x) might not be the way it is formatted in the xlsx file.
        col_width <- vapply(df, function(x) max(nchar(format(x))), NA_real_)

        # message() should be used instead if we really needed to show this
        # print(col_width)

        # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column

        # TODO save this instead as internal package data for quicker loading
        fw <- system.file("extdata", "fontwidth/FontWidth.csv", package = "openxlsx2")
        font_width_tab <- read.csv(fw)

        # TODO base font might not be the font used in this column
        base_font <- wb_get_base_font(self)
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
        col_df <- col_to_df(read_xml(self$createCols(sheet, n = max(cols))))

      # found a few cols, but not all required cols. create the missing columns
      if (any(!cols %in% as.numeric(col_df$min))) {
        beg <- max(as.numeric(col_df$min)) + 1
        end <- max(cols)

        # new columns
        new_cols <- col_to_df(read_xml(self$createCols(sheet, beg = beg, end = end)))

        # rbind only the missing columns. avoiding dups
        sel <- !new_cols$min %in% col_df$min
        col_df <- rbind(col_df, new_cols[sel, ])
        col_df <- col_df[order(as.numeric(col_df[, "min"])), ]
      }

      select <- as.numeric(col_df$min) %in% cols
      col_df$width[select] <- widths
      col_df$hidden[select] <- tolower(hidden)
      col_df$bestFit[select] <- bestFit
      col_df$customWidth[select] <- customWidth
      self$worksheets[[sheet]]$fold_cols(col_df)
      self
    },

    ## rows ----

    # TODO groupRows() and groupCols() are very similiar.  Can problem turn
    # these into some wrappers for another method

    #' @description
    #' Group rows
    #' @param sheet sheet
    #' @param rows rows
    #' @param collapsed collapsed
    #' @param levels levels
    #' @return The `wbWorkbook` object, invisibly
    group_rows = function(sheet, rows, collapsed = FALSE, levels = NULL) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      sheet <- private$get_sheet(sheet)

      if (length(collapsed) > length(rows)) {
        stop("Collapses argument is of greater length than number of rows.")
      }

      if (!is.logical(collapsed)) {
        stop("Collapses should be a logical value (TRUE/FALSE).")
      }

      if (any(rows <= 0L)) {
        stop("Invalid rows entered (<= 0).")
      }

      collapsed <- rep(as.character(as.integer(collapsed)), length.out = length(rows))

      levels <- levels %||% rep("1", length(rows))

      # Remove duplicates
      ok <- !duplicated(rows)
      collapsed <- collapsed[ok]
      levels <- levels[ok]
      rows <- rows[ok]
      sheet <- private$get_sheet(sheet)

      # fetch the row_attr data.frame
      row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

      rows_rev <- rev(rows)

      # get the selection based on the row_attr frame.

      # the first n -1 rows get outlineLevel
      select <- row_attr$r %in% as.character(rows_rev[-1])
      if (length(select)) {
        row_attr$outlineLevel[select] <- as.character(levels[-1])
        row_attr$collapsed[select] <- as.character(as.integer(collapsed[-1]))
        row_attr$hidden[select] <- as.character(as.integer(collapsed[-1]))
      }

      # the n-th row gets only collapsed
      select <- row_attr$r %in% as.character(rows_rev[1])
      if (length(select)) {
        row_attr$collapsed[select] <- as.character(as.integer(collapsed[1]))
      }

      self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr

      # check if there are valid outlineLevel in row_attr and assign outlineLevelRow the max outlineLevel (thats in the documentation)
      if (any(row_attr$outlineLevel != "")) {
        self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelRow = as.character(max(as.integer(row_attr$outlineLevel), na.rm = TRUE))))
      }

      invisible(self)
    },

    #' @description ungroup rows
    #' @param sheet sheet
    #' @param rows rows
    #' @return The `wbWorkbook` object
    ungroup_rows = function(sheet, rows) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      sheet <- private$get_sheet(sheet)

      # check if any rows are selected
      if (any(rows < 1L)) {
        stop("Invalid rows entered (<= 0).")
      }

      # fetch the row_attr data.frame
      row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

      # get the selection based on the row_attr frame.
      select <- row_attr$r %in% as.character(rows)
      if (length(select)) {
        row_attr$outlineLevel[select] <- ""
        row_attr$collapsed[select] <- ""
        # TODO only if unhide = TRUE
        row_attr$hidden[select] <- ""
        self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr
      }

      # If all outlineLevels are missing: remove the outlineLevelRow attribute. Assigning "" will remove the attribute
      if (all(row_attr$outlineLevel == "")) {
        self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelRow = ""))
      } else {
        self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(self$worksheets[[sheet]]$sheetFormatPr, xml_attributes = c(outlineLevelRow = as.character(max(as.integer(row_attr$outlineLevel)))))
      }

      self
    },

    #' @description
    #' Remove a worksheet
    #' @param sheet The worksheet to delete
    #' @return The `wbWorkbook` object, invisibly
    remove_worksheet = function(sheet) {
      # To delete a worksheet
      # Remove colwidths element
      # Remove drawing partname from Content_Types (drawing(sheet).xml)
      # Remove highest sheet from Content_Types
      # Remove drawings element
      # Remove drawings_rels element

      # Remove vml element
      # Remove vml_rels element

      # Remove rowHeights element
      # Remove last sheet element from workbook
      # Remove last sheet element from workbook.xml.rels
      # Remove element from worksheets
      # Remove element from worksheets_rels
      # Remove hyperlinks
      # Reduce calcChain i attributes & remove calcs on sheet
      # Remove sheet from sheetOrder
      # Remove queryTable references from workbook$definedNames to worksheet
      # remove tables

      # TODO can we allow multiple sheets?
      if (length(sheet) != 1) {
        stop("sheet must have length 1.")
      }

      sheet       <- private$get_sheet(sheet)
      sheet_names <- self$sheet_names
      nSheets     <- length(sheet_names)
      sheet_names <- sheet_names[[sheet]]

      ## definedNames
      if (length(self$workbook$definedNames)) {
        # wb_validate_sheet() makes sheet an integer
        # so we need to remove this before getting rid of the sheet names
        self$workbook$definedNames <- self$workbook$definedNames[!get_named_regions(self)$sheets %in% self$sheet_names[sheet]]
      }

      self$remove_named_region(sheet)
      self$sheet_names <- self$sheet_names[-sheet]

      xml_rels <- rbindlist(
         xml_attr(self$worksheets_rels[[sheet]], "Relationship")
      )

      if (nrow(xml_rels)) {
        xml_rels$type   <- basename(xml_rels$Type)
        xml_rels$target <- basename(xml_rels$Target)
        xml_rels$target[xml_rels$type == "hyperlink"] <- ""
        xml_rels$target_ind <- as.numeric(gsub("\\D+", "", xml_rels$target))
      }

      comment_id    <- xml_rels$target_ind[xml_rels$type == "comments"]
      drawing_id    <- xml_rels$target_ind[xml_rels$type == "drawing"]
      pivotTable_id <- xml_rels$target_ind[xml_rels$type == "pivotTable"]
      table_id      <- xml_rels$target_ind[xml_rels$type == "table"]
      thrComment_id <- xml_rels$target_ind[xml_rels$type == "threadedComment"]
      vmlDrawing_id <- xml_rels$target_ind[xml_rels$type == "vmlDrawing"]

      # NULL the sheets
      if (length(comment_id))    self$comments[[comment_id]]            <- NULL
      if (length(drawing_id))    self$drawings[[drawing_id]]            <- NULL
      if (length(drawing_id))    self$drawings_rels[[drawing_id]]       <- NULL
      if (length(thrComment_id)) self$threadComments[[thrComment_id]]   <- NULL
      if (length(vmlDrawing_id)) self$vml[[vmlDrawing_id]]              <- NULL
      if (length(vmlDrawing_id)) self$vml_rels[[vmlDrawing_id]]         <- NULL

      self$isChartSheet <- self$isChartSheet[-sheet]

      #### Modify Content_Types
      ## remove last drawings(sheet).xml from Content_Types
      drawing_name <- xml_rels$target[xml_rels$type == "drawing"]
      if (!is.null(drawing_name)) self$Content_Types <- grep(drawing_name, self$Content_Types, invert = TRUE, value = TRUE)

      ## remove highest sheet
      # (don't chagne this to a "grep(value = TRUE)" ... )
      self$Content_Types <- self$Content_Types[!grepl(sprintf("sheet%s.xml", nSheets), self$Content_Types)]

      # The names for the other drawings have changed
      de <- xml_node(read_xml(self$Content_Types), "Default")
      ct <- rbindlist(xml_attr(read_xml(self$Content_Types), "Override"))
      ct[grepl("drawing", ct$PartName), "PartName"] <- sprintf("/xl/drawings/drawing%i.xml", seq_along(self$drawings))
      ct <- df_to_xml("Override", ct[c("PartName", "ContentType")])
      self$Content_Types <- c(de, ct)


      ## sheetOrder
      toRemove <- which(self$sheetOrder == sheet)
      self$sheetOrder[self$sheetOrder > sheet] <- self$sheetOrder[self$sheetOrder > sheet] - 1L
      self$sheetOrder <- self$sheetOrder[-toRemove]

      # TODO regexpr should be replaced
      ## Need to remove reference from workbook.xml.rels to pivotCache
      removeRels <- grep("pivotTables", self$worksheets_rels[[sheet]], value = TRUE)
      if (length(removeRels)) {
        ## sheet rels links to a pivotTable file, the corresponding pivotTable_rels file links to the cacheDefn which is listing in workbook.xml.rels
        ## remove reference to this file from the workbook.xml.rels
        fileNo <- reg_match0(removeRels, "(?<=pivotTable)[0-9]+(?=\\.xml)")
        fileNo <- as.integer(unlist(fileNo))

        toRemove <- stri_join(
          sprintf("(pivotCacheDefinition%i\\.xml)", fileNo),
          sep = " ",
          collapse = "|"
        )

        toRemove <- stri_join(
          sprintf("(pivotCacheDefinition%i\\.xml)", grep(toRemove, self$pivotTables.xml.rels)),
          sep = " ",
          collapse = "|"
        )

        ## remove reference to file from workbook.xml.res
        self$workbook.xml.rels <- grep(toRemove, self$workbook.xml.rels, invert = TRUE, value = TRUE)
      }

      ## As above for slicers
      ## Need to remove reference from workbook.xml.rels to pivotCache
      # removeRels <- grepl("slicers", self$worksheets_rels[[sheet]])

      if (any(grepl("slicers", self$worksheets_rels[[sheet]]))) {
        # don't change to a grep(value = TRUE)
        self$workbook.xml.rels <- self$workbook.xml.rels[!grepl(sprintf("(slicerCache%s\\.xml)", sheet), self$workbook.xml.rels)]
      }

      ## wont't remove tables and then won't need to reassign table r:id's but will rename them!
      self$worksheets[[sheet]] <- NULL
      self$worksheets_rels[[sheet]] <- NULL

      # tableName is a character Vector with an attached name Vector.
      if (length(self$tables)) {
        nams <- names(self$tables)
        self$tables[table_id] <- ""
        nams[table_id] <- ""
        names(self$tables) <- nams

        tab_sheet <- attr(self$tables, "sheet")
        tab_sheet[table_id] <- 0
        tab_sheet[tab_sheet > sheet] <- tab_sheet[tab_sheet > sheet] - 1L
        attr(self$tables, "sheet") <- tab_sheet

        tab_name <- attr(self$tables, "tableName")
        tab_name[table_id] <- paste0(tab_name[table_id], "_openxlsx_deleted")
        attr(self$tables, "tableName") <- tab_name
      }

      ## drawing will always be the first relationship
      if (nSheets > 1) {
        for (i in seq_len(nSheets - 1L)) {
          # did this get updated from length of 3 to 2?
          #self$worksheets_rels[[i]][1:2] <- genBaseSheetRels(i)
          rel <- rbindlist(xml_attr(self$worksheets_rels[[i]], "Relationship"))
          if (nrow(rel)) {
            if (any(basename(rel$Type) == "drawing")) {
              rel$Target[basename(rel$Type) == "drawing"] <- sprintf("../drawings/drawing%s.xml", i)
            }
            if (is.null(rel$TargetMode)) rel$TargetMode <- ""
            self$worksheets_rels[[i]] <- df_to_xml("Relationship", rel[c("Id", "Type", "Target", "TargetMode")])
          }
        }
      } else {
        self$worksheets_rels <- list()
      }

      ## remove sheet
      sn <- apply_reg_match0(self$workbook$sheets, pat = '(?<= name=")[^"]+')
      self$workbook$sheets <- self$workbook$sheets[!sn %in% sheet_names]

      ## Reset rIds
      if (nSheets > 1) {
        for (i in (sheet + 1L):nSheets) {
          self$workbook$sheets <- gsub(
            stri_join("rId", i),
            stri_join("rId", i - 1L),
            self$workbook$sheets,
            fixed = TRUE
          )
        }
      } else {
        self$workbook$sheets <- NULL
      }

      ## Can remove highest sheet
      # (don't use grepl(value = TRUE))
      self$workbook.xml.rels <- self$workbook.xml.rels[!grepl(sprintf("sheet%s.xml", nSheets), self$workbook.xml.rels)]

      invisible(self)
    },

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
    #' @returns The `wbWorkbook` object
    add_data_validation = function(
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
      if ((type == "date") && !inherits(value, "Date")) {
        stop("If type == 'date' value argument must be a Date vector.")
      }

      if ((type == "time") && !inherits(value, c("POSIXct", "POSIXt"))) {
        stop("If type == 'date' value argument must be a POSIXct or POSIXlt vector.")
      }


      value <- head(value, 2)
      allowBlank <- as.integer(allowBlank[1])
      showInputMsg <- as.integer(showInputMsg[1])
      showErrorMsg <- as.integer(showErrorMsg[1])

      if (type == "list") {
        private$data_validation_list(
          sheet        = sheet,
          startRow     = min(rows),
          endRow       = max(rows),
          startCol     = min(cols),
          endCol       = max(cols),
          value        = value,
          allowBlank   = allowBlank,
          showInputMsg = showInputMsg,
          showErrorMsg = showErrorMsg
        )
      } else {
        private$data_validation(
          sheet        = sheet,
          startRow     = min(rows),
          endRow       = max(rows),
          startCol     = min(cols),
          endCol       = max(cols),
          type         = type,
          operator     = operator,
          value        = value,
          allowBlank   = allowBlank,
          showInputMsg = showInputMsg,
          showErrorMsg = showErrorMsg
        )
      }

      self
    },

    #' @description
    #' Set conditional formatting for a sheet
    #' @param sheet sheet
    #' @param startRow startRow
    #' @param endRow endRow
    #' @param startCol startCol
    #' @param endCol endCol
    #' @param dxfId dxfId
    #' @param formula formula
    #' @param type type
    #' @param values values
    #' @param params params
    #' @return The `wbWorkbook` object, invisibly
    conditional_formatting = function(
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
      sheet <- private$get_sheet(sheet)
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


      # big switch statement
      cfRule <- switch(
        type,

        ## colourScale ----
        colorScale = {

          ## formula contains the colours
          ## values contains numerics or is NULL
          ## dxfId is ignored

          if (is.null(values)) {
          # could use a switch() here for length to also check against other
          # lengths, if these aren't checked somewhere already?
            if (length(formula) == 2L) {
              sprintf(
                # TODO is this indentation necessary?
                '<cfRule type="colorScale" priority="1"><colorScale>
                             <cfvo type="min"/><cfvo type="max"/>
                             <color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>',
                formula[[1]],
                formula[[2]]
              )
            } else {
              sprintf(
                '<cfRule type="colorScale" priority="1"><colorScale>
                             <cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>
                             <color rgb="%s"/><color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>',
                formula[[1]],
                formula[[2]],
                formula[[3]]
              )
            }
          } else {
            if (length(formula) == 2L) {
              sprintf(
                '<cfRule type="colorScale" priority="1"><colorScale>
                            <cfvo type="num" val="%s"/><cfvo type="num" val="%s"/>
                            <color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>',
                values[[1]],
                values[[2]],
                formula[[1]],
                formula[[2]]
              )
            } else {
              sprintf(
                '<cfRule type="colorScale" priority="1"><colorScale>
                            <cfvo type="num" val="%s"/><cfvo type="num" val="%s"/><cfvo type="num" val="%s"/>
                            <color rgb="%s"/><color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>',
                values[[1]],
                values[[2]],
                values[[3]],
                formula[[1]],
                formula[[2]],
                formula[[3]]
              )
            }
          }
        },

        ## dataBar ----
        dataBar = {
          if (length(formula) == 2L) {
            negColour <- formula[[1]]
            posColour <- formula[[2]]
          } else {
            posColour <- formula
            negColour <- "FFFF0000"
          }

          guid <- stri_join(
            "F7189283-14F7-4DE0-9601-54DE9DB",
            40000L + length(self$worksheets[[sheet]]$extLst)
          )

          showValue <- as.integer(params$showValue %||% 1L)
          gradient  <- as.integer(params$gradient  %||% 1L)
          border    <- as.integer(params$border    %||% 1L)

          private$append_sheet_field(sheet, "extLst", {
            gen_databar_extlst(
              guid      = guid,
              sqref     = sqref,
              posColour = posColour,
              negColour = negColour,
              values    = values,
              border    = border,
              gradient  = gradient
            )
          })

          if (is.null(values)) {
            sprintf(
              '<cfRule type="dataBar" priority="1"><dataBar showValue="%s">
                          <cfvo type="min"/><cfvo type="max"/>
                          <color rgb="%s"/>
                          </dataBar>
                          <extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{%s}</x14:id></ext>
                        </extLst></cfRule>',
              showValue,
              posColour,
              guid
            )
          } else {
            sprintf(
              '<cfRule type="dataBar" priority="1"><dataBar showValue="%s">
                            <cfvo type="num" val="%s"/><cfvo type="num" val="%s"/>
                            <color rgb="%s"/>
                            </dataBar>
                            <extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
                            <x14:id>{%s}</x14:id></ext></extLst></cfRule>',
              showValue,
              values[[1]],
              values[[2]],
              posColour,
              guid
            )
          }
        },

        ## expression ----
        expression = {
          sprintf(
            '<cfRule type="expression" dxfId="%s" priority="1"><formula>%s</formula></cfRule>',
            dxfId,
            formula
          )
        },

        ## duplicatedValues ----
        duplicatedValues = {
          sprintf(
            '<cfRule type="duplicateValues" dxfId="%s" priority="1"/>',
            dxfId
          )
        },

        ## containsText ----
        containsText = {
          sprintf(
            '<cfRule type="containsText" dxfId="%s" priority="1" operator="containsText" text="%s">
                        	<formula>NOT(ISERROR(SEARCH("%s", %s)))</formula>
                       </cfRule>',
            dxfId,
            values,
            values,
            # is this unlist correct?  Would this not work?
            # > strsplit(sqref, split = ":")[[1]]
            unlist(strsplit(sqref, split = ":"))[[1]]
          )
        },

        ## notContainsText ----
        notContainsText = {
          sprintf(
            '<cfRule type="notContainsText" dxfId="%s" priority="1" operator="notContains" text="%s">
                        	<formula>ISERROR(SEARCH("%s", %s))</formula>
                       </cfRule>',
            dxfId,
            values,
            values,
            unlist(strsplit(sqref, split = ":"))[[1]]
          )
        },

        ## beginsWith ----
        beginsWith = {
          sprintf(
            '<cfRule type="beginsWith" dxfId="%s" priority="1" operator="beginsWith" text="%s">
                        	<formula>LEFT(%s,LEN("%s"))="%s"</formula>
                       </cfRule>',
            dxfId,
            values,

            unlist(strsplit(sqref, split = ":"))[[1]],
            values,
            values
          )
        },

        ## endsWith ----
        endsWith = sprintf(
          '<cfRule type="endsWith" dxfId="%s" priority="1" operator="endsWith" text="%s">
                        	<formula>RIGHT(%s,LEN("%s"))="%s"</formula>
                       </cfRule>',
          dxfId,
          values,

          unlist(strsplit(sqref, split = ":"))[[1]],
          values,
          values
        ),

        ## between ----
        between = sprintf(
          '<cfRule type="cellIs" dxfId="%s" priority="1" operator="between"><formula>%s</formula><formula>%s</formula></cfRule>',
          dxfId,
          formula[1],
          formula[2]
        ),

        ## topN ----
        topN = sprintf(
          '<cfRule type="top10" dxfId="%s" priority="1" rank="%s" percent="%s"></cfRule>',
          dxfId,
          values[1],
          values[2]
        ),

        ## bottomN ----
        bottomN = {
          sprintf(
            '<cfRule type="top10" dxfId="%s" priority="1" rank="%s" percent="%s" bottom="1"></cfRule>',
            dxfId,
            values[1],
            values[2]
          )
        },
        # do we have a match.arg() anywhere or will it just be showned in this switch()?
        stop("type `", type, "` is not a valid formatting rule")
      )

      private$append_sheet_field(sheet, "conditionalFormatting", cfRule)
      names(self$worksheets[[sheet]]$conditionalFormatting) <- nms
      invisible(self)
    },

    #' @description
    #' Set cell merging for a sheet
    #' @param sheet sheet
    #' @param rows,cols Row and column specifications.
    #' @return The `wbWorkbook` object, invisibly
    merge_cells = function(sheet, rows = NULL, cols = NULL) {
      sheet <- private$get_sheet(sheet)

      # TODO send to wbWorksheet() method
      # self$worksheets[[sheet]]$merge_cells(rows = rows, cols = cols)
      # invisible(self)

      rows <- range(as.integer(rows))
      cols <- range(as.integer(cols))

      # sqref <- get_cell_refs(data.frame(x = rows, y = cols))
      sqref <- paste0(int2col(cols), rows)

      # TODO If the cell merge specs were saved as a data.frame or matrix
      # this would be quicker to check
      current <- reg_match0(self$worksheets[[sheet]]$mergeCells, "[A-Z0-9]+:[A-Z0-9]+")

      # regmatch0 will return character(0) when x is NULL
      if (length(current)) {
        comps <- lapply(
          current,
          function(rectCoords) {
            unlist(strsplit(rectCoords, split = ":"))
          }
        )

        current_cells <- build_cell_merges(comps = comps)
        new_merge <- unlist(build_cell_merges(comps = list(sqref))) # used below in vapply()
        intersects <- vapply(current_cells, function(x) any(x %in% new_merge), NA)

        # Error if merge intersects
        if (any(intersects)) {
          msg <- sprintf(
            "Merge intersects with existing merged cells: \n\t\t%s.\nRemove existing merge first.",
            stri_join(current[intersects], collapse = "\n\t\t")
          )
          stop(msg, call. = FALSE)
        }
      }

        # TODO does this have to be xml?  Can we just save the data.frame or
        # matrix and then check that?  This would also simplify removing the
        # merge specifications
      private$append_sheet_field(sheet, "mergeCells", sprintf('<mergeCell ref="%s"/>', stri_join(sqref, collapse = ":", sep = " " )))
      invisible(self)
    },

    #' @description
    #' Removes cell merging for a sheet
    #' @param sheet sheet
    #' @param rows,cols Row and column specifications.
    #' @return The `wbWorkbook` object, invisibly
    unmerge_cells = function(sheet, rows = NULL, cols = NULL) {
      sheet <- private$get_sheet(sheet)
      rows <- range(as.integer(rows))
      cols <- range(as.integer(cols))
      # sqref <- get_cell_refs(data.frame(x = rows, y = cols))
      sqref <- paste0(int2col(cols), rows)

      current <- regmatches(
        self$worksheets[[sheet]]$mergeCells,
        regexpr("[A-Z0-9]+:[A-Z0-9]+", self$worksheets[[sheet]]$mergeCells)
      )

      if (!is.null(current)) {
        comps <- lapply(current, function(x) unlist(strsplit(x, split = ":")))
        current_cells <- build_cell_merges(comps = comps)
        new <- unlist(build_cell_merges(comps = list(sqref))) # used right below
        mergeIntersections <- vapply(current_cells, function(x) any(x %in% new), NA)

        # Remove intersection
        self$worksheets[[sheet]]$mergeCells <- self$worksheets[[sheet]]$mergeCells[!mergeIntersections]
      }

      invisible(self)
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
      sheet,
      firstActiveRow = NULL,
      firstActiveCol = NULL,
      firstRow = FALSE,
      firstCol = FALSE
    ) {
      # TODO rename to setFreezePanes?
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      # fine to do the validation before the actual check to prevent other errors
      sheet <- private$get_sheet(sheet)

      if (is.null(firstActiveRow) & is.null(firstActiveCol) & !firstRow & !firstCol) {
        return(invisible(self))
      }

      # TODO simplify asserts
      if (!is.logical(firstRow)) stop("firstRow must be TRUE/FALSE")
      if (!is.logical(firstCol)) stop("firstCol must be TRUE/FALSE")

      # make overwrides for arguments
      if (firstRow & !firstCol) {
        firstActiveCol <- NULL
        firstActiveRow <- NULL
        firstCol <- FALSE
      } else if (firstCol & !firstRow) {
        firstActiveRow <- NULL
        firstActiveCol <- NULL
        firstRow <- FALSE
      } else if (firstRow & firstCol) {
        firstActiveRow <- 2L
        firstActiveCol <- 2L
        firstRow <- FALSE
        firstCol <- FALSE
      } else {
        ## else both firstRow and firstCol are FALSE
        firstActiveRow <- firstActiveRow %||% 1L
        firstActiveCol <- firstActiveCol %||% 1L

        # Convert to numeric if column letter given
        # TODO is col2int() safe for non characters?
        firstActiveRow <- col2int(firstActiveRow)
        firstActiveCol <- col2int(firstActiveCol)
      }

      paneNode <-
        if (firstRow) {
          '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        } else if (firstCol) {
          '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>'
        } else {
          if (firstActiveRow == 1 & firstActiveCol == 1) {
            ## nothing to do
            # return(NULL)
            return(invisible(self))
          }

          if (firstActiveRow > 1 & firstActiveCol == 1) {
            attrs <- sprintf('ySplit="%s"', firstActiveRow - 1L)
            activePane <- "bottomLeft"
          }

          if (firstActiveRow == 1 & firstActiveCol > 1) {
            attrs <- sprintf('xSplit="%s"', firstActiveCol - 1L)
            activePane <- "topRight"
          }

          if (firstActiveRow > 1 & firstActiveCol > 1) {
            attrs <- sprintf('ySplit="%s" xSplit="%s"',
              firstActiveRow - 1L,
              firstActiveCol - 1L
            )
            activePane <- "bottomRight"
          }

          sprintf(
            '<pane %s topLeftCell="%s" activePane="%s" state="frozen"/><selection pane="%s"/>',
            stri_join(attrs, collapse = " ", sep = " "),
            get_cell_refs(data.frame(firstActiveRow, firstActiveCol)),
            activePane,
            activePane
          )
        }

      self$worksheets[[sheet]]$freezePane <- paneNode
      invisible(self)
    },

    ## plots and images ----

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
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      if (!file.exists(file)) {
        stop("File does not exist.")
      }

      # TODO require user to pass a valid path
      if (!grepl("\\\\|\\/", file)) {
        file <- file.path(getwd(), file, fsep = .Platform$file.sep)
      }

      units <- tolower(units)

      # TODO use match.arg()
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
      width <- as.integer(round(width * 914400L, 0)) # (EMUs per inch)
      height <- as.integer(round(height * 914400L, 0)) # (EMUs per inch)

      ## within the sheet the drawing node's Id refernce an id in the sheetRels
      ## sheet rels reference the drawingi.xml file
      ## drawingi.xml refernece drawingRels
      ## drawing rels reference an image in the media folder
      ## worksheetRels(sheet(i)) references drawings(j)

      sheet <- private$get_sheet(sheet)

      # TODO tools::file_ext() ...
      imageType <- regmatches(file, gregexpr("\\.[a-zA-Z]*$", file))
      imageType <- gsub("^\\.", "", imageType)

      drawing_len <- 0
      if (length(self$drawings_rels[[sheet]]))
        drawing_len <- length(xml_node(unlist(self$drawings_rels[[sheet]]), "Relationship"))

      imageNo <- drawing_len + 1L
      mediaNo <- length(self$media) + 1L

      startCol <- col2int(startCol)

      ## update Content_Types
      if (!any(grepl(stri_join("image/", imageType), self$Content_Types))) {
        self$Content_Types <-
          unique(c(
            sprintf(
              '<Default Extension="%s" ContentType="image/%s"/>',
              imageType,
              imageType
            ),
            self$Content_Types
          ))
      }

      ## drawings rels (Reference from drawings.xml to image file in media folder)
      self$drawings_rels[[sheet]] <- paste0(
        unlist(self$drawings_rels[[sheet]]),
        sprintf(
          '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%s.%s"/>',
          imageNo,
          mediaNo,
          imageType
        )
      )

      ## write file path to media slot to copy across on save
      tmp <- file
      names(tmp) <- stri_join("image", mediaNo, ".", imageType)
      self$append("media", tmp)

      ## create drawing.xml
      anchor <- '<xdr:oneCellAnchor>'

      from <- sprintf(
        '<xdr:from>
        <xdr:col>%s</xdr:col>
        <xdr:colOff>%s</xdr:colOff>
        <xdr:row>%s</xdr:row>
        <xdr:rowOff>%s</xdr:rowOff>
        </xdr:from>',
        startCol - 1L,
        colOffset,
        startRow - 1L,
        rowOffset
      )

      drawingsXML <- stri_join(
        anchor,
        from,
        sprintf('<xdr:ext cx="%s" cy="%s"/>', width, height),
        genBasePic(imageNo),
        "<xdr:clientData/>",
        "</xdr:oneCellAnchor>"
      )


      # If no drawing is found, initiate one. If one is found, append a child to the exisiting node.
      # Might look into updating attributes as well.
      if (length(self$drawings[[sheet]]) == 0) {
        xml_attr = c(
          "xmlns:xdr" = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
          "xmlns:a" = "http://schemas.openxmlformats.org/drawingml/2006/main",
          "xmlns:r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        )
        self$drawings[[sheet]] <- xml_node_create("xdr:wsDr", xml_children = drawingsXML, xml_attributes = xml_attr)
      } else {
        self$drawings[[sheet]] <- xml_add_child(self$drawings[[sheet]], drawingsXML)
      }

      invisible(self)
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
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      if (is.null(dev.list()[[1]])) {
        warning("No plot to insert.")
        return(self)
      }

      if (!is.null(xy)) {
        startCol <- xy[[1]]
        startRow <- xy[[2]]
      }

      fileType <- tolower(fileType)
      units <- tolower(units)

      # TODO just don't allow jpg
      if (fileType == "jpg") {
        fileType <- "jpeg"
      }

      # TODO add match.arg()
      if (!fileType %in% c("png", "jpeg", "tiff", "bmp")) {
        stop("Invalid file type.\nfileType must be one of: png, jpeg, tiff, bmp")
      }

      if (!units %in% c("cm", "in", "px")) {
        stop("Invalid units.\nunits must be one of: cm, in, px")
      }

      fileName <- tempfile(pattern = "figureImage", fileext = paste0(".", fileType))

      # TODO use switch()
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

      self$add_image(
        sheet     = sheet,
        file      = fileName,
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


    #' @description
    #' Prints the `wbWorkbook` object
    #' @return The `wbWorkbook` object, invisibly; called for its side-effects
    print = function() {
      exSheets <- self$sheet_names
      nSheets <- length(exSheets)
      nImages <- length(self$media)
      nCharts <- length(self$charts)

      exSheets <- replaceXMLEntities(exSheets)
      showText <- "A Workbook object.\n"

      ## worksheets
      if (nSheets > 0) {
        showText <- c(showText, "\nWorksheets:\n")

        # TODO use seq_along()
        sheetTxt <- lapply(seq_len(nSheets), function(i) {
          tmpTxt <- sprintf('Sheet %s: "%s"\n', i, exSheets[[i]])

          if (length(self$rowHeights[[i]])) {
            tmpTxt <-
              c(
                tmpTxt,
                c(
                  "\n\tCustom row heights (row: height)\n\t",
                  stri_join(
                    sprintf("%s: %s", names(self$rowHeights[[i]]), round(as.numeric(
                      self$rowHeights[[i]]
                    ), 2)),
                    collapse = ", ",
                    sep = " "
                  )
                )
              )
          }
        })

        showText <- c(showText, sheetTxt, "\n")
      } else {
        showText <-
          c(showText, "\nWorksheets:\n", "No worksheets attached\n")
      }

      ## images
      if (nImages > 0) {
        showText <-
          c(
            showText,
            "\nImages:\n",
            sprintf('Image %s: "%s"\n', seq_len(nImages), self$media)
          )
      }

      if (nCharts > 0) {
        showText <-
          c(
            showText,
            "\nCharts:\n",
            sprintf('Chart %s: "%s"\n', seq_len(nCharts), self$charts)
          )
      }

      if (nSheets > 0) {
        showText <-
          c(showText, sprintf(
            "Worksheet write order: %s",
            stri_join(self$sheetOrder, sep = " ", collapse = ", ")
          ))
      }

      cat(unlist(showText))
      invisible(self)
    },


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
      protect = TRUE,
      lockStructure = FALSE,
      lockWindows = FALSE,
      password = NULL,
      type = NULL,
      fileSharing = FALSE,
      username = unname(Sys.info()["user"]),
      readOnlyRecommended = FALSE
    ) {

      attr <- vector("character", 3L)
      names(attr) <- c("workbookPassword", "lockStructure", "lockWindows")

      if (!is.null(password)) {
        attr["workbookPassword"] <- hashPassword(password)
      }
      if (!missing(lockStructure) && !is.null(lockStructure)) {
        attr["lockStructure"] <- toString(as.numeric(lockStructure))
      }
      if (!missing(lockWindows) && !is.null(lockWindows)) {
        attr["lockWindows"] <- toString(as.numeric(lockWindows))
      }

      # TODO: Shall we parse the existing protection settings and preserve all unchanged attributes?
      if (protect) {
        self$workbook$workbookProtection <-
          xml_node_create("workbookProtection", xml_attributes = attr[attr != ""])

        # TODO: use xml_node_create
        if (fileSharing) {
          if (type == 2L) readOnlyRecommended <- TRUE
          fileSharingPassword <- function(x, username, readOnlyRecommended) {
            readonly <- ifelse(readOnlyRecommended, 'readOnlyRecommended="1"', '')
            sprintf('<fileSharing userName="%s" %s reservationPassword="%s"/>', username, readonly, x)
          }

          self$workbook$fileSharing <- fileSharingPassword(attr["workbookPassword"], username, readOnlyRecommended)
        }

        if (!is.null(type) | !is.null(password))
          self$workbook$apps <- sprintf("<DocSecurity>%i</DocSecurity>", type)

      } else {
        self$workbook$workbookProtection <- ""
      }

      invisible(self)
    },


    ### creators --------------------------------------------------------------

    #' @description Set creator(s)
    #' @param creators A character vector of creators to set.  Duplicates are
    #'   ignored.
    set_creators = function(creators) {
      private$modify_creators("set", creators)
    },


    #' @description Add creator(s)
    #' @param creators A character vector of creators to add.  Duplicates are
    #'   ignored.
    add_creators = function(creators) {
      private$modify_creators("add", creators)
    },


    #' @description Remove creator(s)
    #' @param creators A character vector of creators to remove.  All duplicated
    #'   are removed.
    remove_creators = function(creators) {
      private$modify_creators("remove", creators)
    },


    #' @description
    #' Change the last modified by
    #' @param LastModifiedBy A new value
    #' @return The `wbWorkbook` object, invisibly
    set_last_modified_by = function(LastModifiedBy = NULL) {
      # TODO rename to wb_set_last_modified_by() ?
      if (!is.null(LastModifiedBy)) {
        current_LastModifiedBy <-
          stri_match(self$core, regex = "<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>")[1, 2]
        self$core <-
          stri_replace_all_fixed(
            self$core,
            pattern = current_LastModifiedBy,
            replacement = LastModifiedBy
          )
      }

      invisible(self)
    },

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
      sheet <- private$get_sheet(sheet)
      xml <- self$worksheets[[sheet]]$pageSetup

      if (!is.null(orientation)) {
        orientation <- tolower(orientation)
        if (!orientation %in% c("portrait", "landscape")) stop("Invalid page orientation.")
      } else {
        # if length(xml) == 1 then use if () {} else {}
        orientation <- ifelse(grepl("landscape", xml), "landscape", "portrait") ## get existing
      }

      if ((scale < 10) || (scale > 400)) {
        stop("Scale must be between 10 and 400.")
      }

      if (!is.null(paperSize)) {
        paperSizes <- 1:68
        paperSizes <- paperSizes[!paperSizes %in% 48:49]
        if (!paperSize %in% paperSizes) {
          stop("paperSize must be an integer in range [1, 68]. See ?ws_page_setup details.")
        }
        paperSize <- as.integer(paperSize)
      } else {
        paperSize <- regmatches(xml, regexpr('(?<=paperSize=")[0-9]+', xml, perl = TRUE)) ## get existing
      }

      ## Keep defaults on orientation, hdpi, vdpi, paperSize ----
      hdpi <- regmatches(xml, regexpr('(?<=horizontalDpi=")[0-9]+', xml, perl = TRUE))
      vdpi <- regmatches(xml, regexpr('(?<=verticalDpi=")[0-9]+', xml, perl = TRUE))

      ## Update ----
      self$worksheets[[sheet]]$pageSetup <- sprintf(
        '<pageSetup paperSize="%s" orientation="%s" scale = "%s" fitToWidth="%s" fitToHeight="%s" horizontalDpi="%s" verticalDpi="%s" r:id="rId2"/>',
        paperSize, orientation, scale, as.integer(fitToWidth), as.integer(fitToHeight), hdpi, vdpi
      )

      if (fitToHeight || fitToWidth) {
        self$worksheets[[sheet]]$sheetPr <- unique(c(self$worksheets[[sheet]]$sheetPr, '<pageSetupPr fitToPage="1"/>'))
      }

      self$worksheets[[sheet]]$pageMargins <-
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
        self$worksheets[[sheet]]$sheetPr <- unique(c(self$worksheets[[sheet]]$sheetPr, paste0("<outlinePr", outlinepr, "/>")))
      }

      ## print Titles ----
      if (!is.null(printTitleRows) && is.null(printTitleCols)) {
        if (!is.numeric(printTitleRows)) {
          stop("printTitleRows must be numeric.")
        }

        private$create_named_region(
          ref1 = paste0("$", min(printTitleRows)),
          ref2 = paste0("$", max(printTitleRows)),
          name = "_xlnm.Print_Titles",
          sheet = names(self)[[sheet]],
          localSheetId = sheet - 1L
        )
      } else if (!is.null(printTitleCols) && is.null(printTitleRows)) {
        if (!is.numeric(printTitleCols)) {
          stop("printTitleCols must be numeric.")
        }

        cols <- int2col(range(printTitleCols))
        private$create_named_region(
          ref1 = paste0("$", cols[1]),
          ref2 = paste0("$", cols[2]),
          name = "_xlnm.Print_Titles",
          sheet = names(self)[[sheet]],
          localSheetId = sheet - 1L
        )
      } else if (!is.null(printTitleCols) && !is.null(printTitleRows)) {
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
        sheet <- names(self)[[sheet]]

        self$workbook$definedNames <- c(
          self$workbook$definedNames,
          sprintf('<definedName name="_xlnm.Print_Titles" localSheetId="%s">\'%s\'!%s,\'%s\'!%s</definedName>', localSheetId, sheet, cols, sheet, rows)
        )

      }

      self
    },

    ## header footer ----

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
      sheet,
      header      = NULL,
      footer      = NULL,
      evenHeader  = NULL,
      evenFooter  = NULL,
      firstHeader = NULL,
      firstFooter = NULL
    ) {
      sheet <- private$get_sheet(sheet)

      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      if (!is.null(header) && length(header) != 3) {
        stop("header must have length 3 where elements correspond to positions: left, center, right.")
      }

      if (!is.null(footer) && length(footer) != 3) {
        stop("footer must have length 3 where elements correspond to positions: left, center, right.")
      }

      if (!is.null(evenHeader) && length(evenHeader) != 3) {
        stop("evenHeader must have length 3 where elements correspond to positions: left, center, right.")
      }

      if (!is.null(evenFooter) && length(evenFooter) != 3) {
        stop("evenFooter must have length 3 where elements correspond to positions: left, center, right.")
      }

      if (!is.null(firstHeader) && length(firstHeader) != 3) {
        stop("firstHeader must have length 3 where elements correspond to positions: left, center, right.")
      }

      if (!is.null(firstFooter) && length(firstFooter) != 3) {
        stop("firstFooter must have length 3 where elements correspond to positions: left, center, right.")
      }

      # TODO this could probably be moved to the hf assignment
      oddHeader   <- headerFooterSub(header)
      oddFooter   <- headerFooterSub(footer)
      evenHeader  <- headerFooterSub(evenHeader)
      evenFooter  <- headerFooterSub(evenFooter)
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

      self$worksheets[[sheet]]$headerFooter <- hf
      self
    },

    #' @description get tables
    #' @param sheet sheet
    #' @returns The sheet tables.  `character()` if empty
    get_tables = function(sheet) {
      if (length(sheet) != 1) {
        stop("sheet argument must be length 1")
      }

      if (length(self$tables) == 0) {
        return(character())
      }

      sheet <- private$get_sheet(sheet)
      if (is.na(sheet)) stop("No such sheet in workbook")

      table_sheets <- attr(self$tables, "sheet")
      tables <- attr(self$tables, "tableName")
      refs <- names(self$tables)

      refs <- refs[table_sheets == sheet & !grepl("openxlsx_deleted", tables, fixed = TRUE)]
      tables <- tables[table_sheets == sheet & !grepl("openxlsx_deleted", tables, fixed = TRUE)]

      if (length(tables)) {
        attr(tables, "refs") <- refs
      }

      return(tables)
    },


    #' @description remove tables
    #' @param sheet sheet
    #' @param table table
    #' @returns The `wbWorkbook` object
    remove_tables = function(sheet, table) {
      if (length(table) != 1) {
        stop("table argument must be length 1")
      }

      ## delete table object and all data in it
      sheet <- private$get_sheet(sheet)

      if (!table %in% attr(self$tables, "tableName")) {
        stop(sprintf("table '%s' does not exist.", table), call. = FALSE)
      }

      ## get existing tables
      table_sheets <- attr(self$tables, "sheet")
      table_names <- attr(self$tables, "tableName")
      refs <- names(self$tables)

      ## delete table object (by flagging as deleted)
      inds <- which(table_sheets %in% sheet & table_names %in% table)
      table_name_original <- table_names[inds]

      table_names[inds] <- paste0(table_name_original, "_openxlsx_deleted")
      attr(self$tables, "tableName") <- table_names

      ## delete reference from worksheet to table
      worksheet_table_names <- attr(self$worksheets[[sheet]]$tableParts, "tableName")
      to_remove <- which(worksheet_table_names == table_name_original)

      self$worksheets[[sheet]]$tableParts <- self$worksheets[[sheet]]$tableParts[-to_remove]
      attr(self$worksheets[[sheet]]$tableParts, "tableName") <- worksheet_table_names[-to_remove]


      ## Now delete data from the worksheet
      refs <- strsplit(refs[[inds]], split = ":")[[1]]
      rows <- as.integer(gsub("[A-Z]", "", refs))
      rows <- seq(from = rows[1], to = rows[2], by = 1)

      cols <- col2int(refs)
      cols <- seq(from = cols[1], to = cols[2], by = 1)

      ## now delete data
      delete_data(wb = self, sheet = sheet, rows = rows, cols = cols, gridExpand = TRUE)
      self
    },

    #' @description add filters
    #' @param sheet sheet
    #' @param rows rows
    #' @param cols cols
    #' @returns The `wbWorkbook` object
    add_filter = function(sheet, rows, cols) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)
      sheet <- private$get_sheet(sheet)

      if (length(rows) != 1) {
        stop("row must be a numeric of length 1.")
      }

      if (!is.numeric(cols)) {
        cols <- col2int(cols)
      }

      self$worksheets[[sheet]]$autoFilter <- sprintf(
        '<autoFilter ref="%s"/>',
        paste(get_cell_refs(data.frame("x" = c(rows, rows), "y" = c(min(cols), max(cols)))), collapse = ":")
      )

      self
    },

    #' @description remove filters
    #' @param sheet sheet
    #' @returns The `wbWorkbook` object
    remove_filter = function(sheet) {
      for (s in private$get_sheet(sheet)) {
        self$worksheets[[s]]$autoFilter <- character()
      }

      self
    },

    #' @description grid lines
    #' @param sheet sheet
    #' @param show show
    #' @returns The `wbWorkbook` object
    grid_lines = function(sheet, show = FALSE) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      sheet <- private$get_sheet(sheet)

      if (!is.logical(show)) {
        stop("show must be a logical")
      }

      sv <- self$worksheets[[sheet]]$sheetViews
      show <- as.integer(show)
      ## If attribute exists gsub
      if (grepl("showGridLines", sv)) {
        sv <- gsub('showGridLines=".?[^"]', sprintf('showGridLines="%s', show), sv, perl = TRUE)
      } else {
        sv <- gsub("<sheetView ", sprintf('<sheetView showGridLines="%s" ', show), sv)
      }

      self$worksheets[[sheet]]$sheetViews <- sv
      self
    },

    ### named region ----

    #' @description add a named region
    #' @param sheet sheet
    #' @param cols cols
    #' @param rows rows
    #' @param name name
    #' @param localSheetId localSheetId
    #' @param overwrite overwrite
    #' @returns The `wbWorkbook` object
    add_named_region = function(
      sheet,
      cols,
      rows,
      name,
      localSheetId = NULL,
      overwrite = FALSE
    ) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      sheet <- private$get_sheet(sheet)

      if (!is.numeric(rows)) {
        stop("rows argument must be a numeric/integer vector")
      }

      if (!is.numeric(cols)) {
        stop("cols argument must be a numeric/integer vector")
      }

      ## check name doesn't already exist
      ## named region

      # TODO use reg_match0?
      ex_names <- regmatches(self$workbook$definedNames, regexpr('(?<=name=")[^"]+', self$workbook$definedNames, perl = TRUE))
      ex_names <- tolower(replaceXMLEntities(ex_names))

      if (tolower(name) %in% ex_names) {
        if (overwrite)
          self$workbook$definedNames <- self$workbook$definedNames[!ex_names %in% tolower(name)]
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

      private$create_named_region(
        ref1         = ref1,
        ref2         = ref2,
        name         = name,
        sheet        = self$sheet_names[sheet],
        localSheetId = localSheetId
      )

      self
    },

    #' @description remove a named region
    #' @param sheet sheet
    #' @param name name
    #' @returns The `wbWorkbook` object
    remove_named_region = function(sheet = NULL, name = NULL) {
      # get all nown defined names
      dn <- get_named_regions(self)

      if (is.null(name) && !is.null(sheet)) {
        sheet <- private$get_sheet(sheet)
        del <- dn$id[dn$sheet == sheet]
      } else if (!is.null(name) && is.null(sheet)) {
        del <- dn$id[dn$name == name]
      } else {
        sheet <- private$get_sheet(sheet)
        del <- dn$id[dn$sheet == sheet & dn$name == name]
      }

      if (length(del)) {
        self$workbook$definedNames <- self$workbook$definedNames[-del]
      } else {
        if (!is.null(name))
          warning(sprintf("Cannot find named region with name '%s'", name))
        # do not warn if wb and sheet are selected. wb_delete_named_region is
        # called with every wb_remove_worksheet and would throw meaningless
        # warnings. For now simply assume if no name is defined, that the
        # user does not care, as long as no defined name remains on a sheet.
      }

      self
    },

    #' @description set worksheet order
    #' @param sheets sheets
    #' @return The `wbWorkbook` object
    set_order = function(sheets) {
      sheets <- private$get_sheet(sheet = sheets)

      if (anyDuplicated(sheets)) {
        stop("`sheets` cannot have duplicates")
      }

      if (length(sheets) != length(self$worksheets)) {
        stop(sprintf("Worksheet order must be same length as number of worksheets [%s]", length(wb$worksheets)))
      }

      if (any(sheets > length(self$worksheets))) {
        stop("Elements of order are greater than the number of worksheets")
      }

      self$sheetOrder <- sheets
      self
    },

    ## sheet visibility ----

    #' @description Get sheet visibility
    #' @returns Returns sheet visibility
    get_sheet_visibility = function() {
      state <- rep("visible", length(self$workbook$sheets))
      state[grepl("hidden", self$workbook$sheets)] <- "hidden"
      state[grepl("veryHidden", self$workbook$sheets, ignore.case = TRUE)] <- "veryHidden"
      state
    },

    #' @description Set sheet visibility
    #' @param value value
    #' @param sheet sheet
    #' @returns The `wbWorkbook` object
    set_sheet_visibility = function(sheet, value) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)

      if (length(value) != length(sheet)) {
        stop("`value` and `sheet` must be the same length")
      }

      sheet <- private$get_sheet(sheet)

      value <- tolower(as.character(value))
      value[value %in% "true"] <- "visible"
      value[value %in% "false"] <- "hidden"
      value[value %in% "veryhidden"] <- "veryHidden"

      exState0 <- reg_match0(self$workbook$sheets[sheet], '(?<=state=")[^"]+')
      exState <- tolower(exState0)
      exState[exState %in% "true"] <- "visible"
      exState[exState %in% "hidden"] <- "hidden"
      exState[exState %in% "false"] <- "hidden"
      exState[exState %in% "veryhidden"] <- "veryHidden"


      inds <- which(value != exState)

      if (length(inds) == 0) {
        return(self)
      }

      for (i in seq_along(self$worksheets)) {
        self$workbook$sheets[sheet[i]] <- gsub(exState0[i], value[i], self$workbook$sheets[sheet[i]], fixed = TRUE)
      }

      if (!any(self$get_sheet_visibility() %in% c("true", "visible"))) {
        warning("A workbook must have atleast 1 visible worksheet.  Setting first for visible")
        self$set_sheet_visibility(1, TRUE)
      }

      self
    },

    ## page breaks ----

    #' @description Add a page break
    #' @param sheet sheet
    #' @param row row
    #' @param col col
    #' @returns The `wbWorkbook` object
    add_page_break = function(sheet, row = NULL, col = NULL) {
      op <- openxlsx_options()
      on.exit(options(op), add = TRUE)
      sheet <- private$get_sheet(sheet)
      self$worksheets[[sheet]]$add_page_break(row = row, col = col)
      self
    }
  ),

  ## private ----

  # any functions that are not present elsewhere or are non-exported internal
  # functions that are used to make assignments
  private = list(
    deep_clone = function(name, value) {
      #' Deep cloning method for workbooks.  This method also accesses
      #' `$clone(deep = TRUE)` methods for `R6` fields.
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

    validate_new_sheet = function(sheet) {
      # returns nothing, throws error if there's a problem.

      if (is.na(sheet)) {
        stop("sheet cannot be NA")
      }

      if (is.numeric(sheet)) {
        if (sheet %% 1 != 0) {
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
      if (any_illegal_chars(sheet)) {
        stop("Illegal characters found in sheet. Please remove. See ?openxlsx::clean_worksheet_name")
      }

      if (nchar(sheet) > 31) {
        stop("sheet names must be <= 32 chars")
      }

      if (tolower(sheet) %in% self$sheet_names) {
        stop("a sheet with name '", sheet, '"already exists"')
      }
    },

    get_sheet = function(sheet) {
      # returns the sheet index, or NA
      if (is.null(self$sheet_names)) {
        warning("Workbook has no sheets")
        return(NA_integer_)
      }

      if (is.character(sheet)) {
        m <- match(tolower(sheet), tolower(self$sheet_names))
        bad <- is.na(m)

        if (any(bad)) {
          stop("Sheet names not found: ", toString(sheet[bad]))
        }

        sheet <- m
      } else {
        sheet <- as.integer(sheet)
        bad <- which(sheet > length(self$sheet_names))

        if (length(bad)) {
          stop("Invalid sheet position(s): ", toString(sheet[bad]))
          # sheet[bad] <- NA_integer_
        }
      }

      sheet
    },

    append_sheet_field = function(sheet, field, value = NULL) {
      # if using this we should consider adding a method into the wbWorksheet
      # object.  wbWorksheet$append() is currently public. _Currently_.
      sheet <- private$get_sheet(sheet)
      self$worksheets[[sheet]]$append(field, value)
      self
    },

    append_workbook_field = function(field, value = NULL) {
      self$workbook[[field]] <- c(self$workbook[[field]], value)
      self
    },

    append_sheet_rels = function(sheet, value = NULL) {
      sheet <- private$get_sheet(sheet)
      self$worksheets_rels[[sheet]] <- c(self$worksheets_rels[[sheet]], value)
      self
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
          if (!is.null(self$title))    sprintf("<dc:title>%s</dc:title>",       replaceIllegalCharacters(self$title)),
          if (!is.null(self$subject))  sprintf("<dc:subject>%s</dc:subject>",   replaceIllegalCharacters(self$subject)),
          if (!is.null(self$category)) sprintf("<cp:category>%s</cp:category>", replaceIllegalCharacters(self$category)),

          # end
          "</coreProperties>",
          collapse = "",
          unlist = TRUE
        )

      invisible(self)
    },

    modify_creators = function(method = c("add", "set", "remove"), value) {
      method <- match.arg(method)
      assert_class(value, "character")

      if (any(!has_chr(value))) {
        stop("all creators must contain characters without NAs", call. = FALSE)
      }

      value <- switch(
        method,
        add    = unique(c(self$creator, value)),
        set    = unique(value),
        remove = setdiff(self$creator, value)
      )

      self$creator <- value
      # core is made on initialization
      private$generate_base_core()
      self
    },

    ws = function(sheet) {
      self$worksheets[[private$get_sheet(sheet)]]
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

    writeDrawingVML = function(dir) {
      for (i in seq_along(self$comments)) {
        id <- 1025

        cd <- unapply(self$comments[[i]], "[[", "clientData")
        nComments <- length(cd)

        ## write head
        if (nComments > 0 | length(self$vml[[i]])) {
          write(
            x = stri_join(
              '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
            ),
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            sep = " "
          )
        }

        # TODO use seq_along()
        for (j in seq_len(nComments)) {
          id <- id + 1L
          write(
            x = genBaseShapeVML(cd[j], id),
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            append = TRUE
          )
        }

        if (length(self$vml[[i]])) {
          write(
            x = self$vml[[i]],
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            append = TRUE
          )
        }

        # TODO nComments and self$vml is already checked
        if (nComments > 0 | length(self$vml[[i]])) {
          write(
            x = "</xml>",
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            append = TRUE
          )
        }

      }

      for (i in seq_along(self$drawings_vml)) {
        write(
          x = self$drawings_vml[[i]],
          file = file.path(dir, sprintf("vmlDrawing%s.vml", i))
        )
      }

      invisible(self)
    },

    writeSheetDataXML = function(
      xldrawingsDir,
      xldrawingsRelsDir,
      xlworksheetsDir,
      xlworksheetsRelsDir
    ) {
      ## write worksheets

      # TODO just seq_along()
      nSheets <- length(self$worksheets)

      for (i in seq_len(nSheets)) {
        ## Write drawing i (will always exist) skip those that are empty
        if (!identical(self$drawings[[i]], list())) {
          write_file(
            head = '',
            body = pxml(self$drawings[[i]]),
            tail = '',
            fl = file.path(xldrawingsDir, stri_join("drawing", i, ".xml"))
          )
          if (!identical(self$drawings_rels[[i]], list())) {
            write_file(
              head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
              body = pxml(self$drawings_rels[[i]]),
              tail = "</Relationships>",
              fl = file.path(xldrawingsRelsDir, stri_join("drawing", i, ".xml.rels"))
            )
          }
        } else {
          self$worksheets[[i]]$drawing <- character()
        }

        ## vml drawing
        if (length(self$vml_rels[[i]])) {
          write_file(
            head = '',
            body = pxml(self$vml_rels[[i]]),
            tail = '',
            fl = file.path(xldrawingsRelsDir, stri_join("vmlDrawing", i, ".vml.rels"))
          )
        }

        if (self$isChartSheet[i]) {
          chartSheetDir <- file.path(dirname(xlworksheetsDir), "chartsheets")
          chartSheetRelsDir <-
            file.path(dirname(xlworksheetsDir), "chartsheets", "_rels")

          if (!file.exists(chartSheetDir)) {
            dir.create(chartSheetDir, recursive = TRUE)
            dir.create(chartSheetRelsDir, recursive = TRUE)
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
            ws$sheet_data$row_attr <- rows_attr[order(as.numeric(rows_attr[, "r"])),]

            cc_rows <- ws$sheet_data$row_attr$r
            cc_out <- cc[cc$row_r %in% cc_rows, c("row_r", "c_r",  "r", "v", "c_t", "c_s", "f", "f_t", "f_ref", "f_ca", "f_si", "is")]

            ws$sheet_data$cc_out <- cc_out[order(as.integer(cc_out[,"row_r"]), col2int(cc_out[, "c_r"])),]
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

          write_worksheet(
            prior = prior,
            post = post,
            sheet_data = ws$sheet_data,
            cols_attr = ws$cols_attr,
            R_fileName = file.path(xlworksheetsDir, sprintf("sheet%s.xml", i))
          )

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
            if (length(self$tables)) {
              table_inds <- grep("tables/table[0-9].xml", ws_rels)

              relship <- rbindlist(xml_attr(ws_rels, "Relationship"))
              relship$typ <- basename(relship$Type)
              relship$tid <- gsub("\\D+", "", relship$Target)

              table_nms <- attr(self$tables, "tableName")

              is_deleted <- which(grepl("_openxlsx_deleted", table_nms, fixed = TRUE))
              delete <- relship$typ == "table" & relship$tid %in% is_deleted

              if (any(delete)) {
                relship <- relship[!delete,]
              }
              relship$typ <- relship$tid <- NULL
              if (is.null(relship$TargetMode)) relship$TargetMode <- ""
              ws_rels <- df_to_xml("Relationship", df_col = relship[c("Id", "Type", "Target", "TargetMode")])
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

      invisible(self)
    },

    data_validation = function(
      sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      type,
      operator,
      value,
      allowBlank,
      showInputMsg,
      showErrorMsg
    ) {
      # TODO rename: setDataValidation?
      # TODO can this be moved to the worksheet class?
      sheet <- private$get_sheet(sheet)
      sqref <-
        stri_join(get_cell_refs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        )),
          sep = " ",
          collapse = ":"
        )

      header <-
        sprintf(
          '<dataValidation type="%s" operator="%s" allowBlank="%s" showInputMessage="%s" showErrorMessage="%s" sqref="%s">',
          type,
          operator,
          allowBlank,
          showInputMsg,
          showErrorMsg,
          sqref
        )


      # TODO consider switch(type, date = ..., time = ..., )
      if (type == "date") {
        # TODO consider origin <- if () ... else ...
        origin <- 25569L
        # TODO would it be faster to just search each self$workbook instead of
        # trying to unlist and join everything?
        if (grepl(
          'date1904="1"|date1904="true"',
          stri_join(unlist(self$workbook), sep = " ", collapse = ""),
          ignore.case = TRUE
        )) {
          origin <- 24107L
        }

        value <- as.integer(value) + origin
      }

      if (type == "time") {
        # TODO simplify with above?  This is the same thing?
        origin <- 25569L
        if (grepl(
          'date1904="1"|date1904="true"',
          stri_join(unlist(self$workbook), sep = " ", collapse = ""),
          ignore.case = TRUE
        )) {
          origin <- 24107L
        }

        t <- format(value[1], "%z")
        offSet <-
          suppressWarnings(
            ifelse(substr(t, 1, 1) == "+", 1L, -1L) * (
              as.integer(substr(t, 2, 3)) + as.integer(substr(t, 4, 5)) / 60
            ) / 24
          )
        if (is.na(offSet)) {
          offSet[i] <- 0
        }

        value <- as.numeric(as.POSIXct(value)) / 86400 + origin + offSet
      }

      form <- sapply(
        seq_along(value),
        function(i) {
          sprintf("<formula%s>%s</formula%s>", i, value[i], i)
        }
      )

      private$append_sheet_field(sheet, "dataValidations", stri_join(header, stri_join(form, collapse = ""), "</dataValidation>"))
      invisible(self)
    },

    data_validation_list = function(
      sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      value,
      allowBlank,
      showInputMsg,
      showErrorMsg
    ) {
      # TODO consider some defaults to logicals
      # TODO rename: setDataValidationList?
      sheet <- private$get_sheet(sheet)
      sqref <-
        stri_join(get_cell_refs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        )),
          sep = " ",
          collapse = ":"
        )
      data_val <-
        sprintf(
          '<x14:dataValidation type="list" allowBlank="%s" showInputMessage="%s" showErrorMessage="%s">',
          allowBlank,
          showInputMsg,
          showErrorMsg
        )

      formula <- sprintf("<x14:formula1><xm:f>%s</xm:f></x14:formula1>", value)
      sqref <- sprintf("<xm:sqref>%s</xm:sqref>", sqref)
      xmlData <- stri_join(data_val, formula, sqref, "</x14:dataValidation>")
      private$append_sheet_field(sheet, "dataValidationsLst", xmlData)
      invisible(self)
    },

    # old add_named_region()
    create_named_region = function(ref1, ref2, name, sheet, localSheetId = NULL) {
      name <- replaceIllegalCharacters(name)
      value <- if (is.null(localSheetId)) {
        sprintf(
          '<definedName name="%s">\'%s\'!%s:%s</definedName>',
          name,
          sheet,
          ref1,
          ref2
        )
      } else {
        sprintf(
          '<definedName name="%s" localSheetId="%s">\'%s\'!%s:%s</definedName>',
          name,
          localSheetId,
          sheet,
          ref1,
          ref2
        )
      }

      private$append_workbook_field("definedNames", value)
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

      sheetRIds <- get_r_id(self)
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
      extRefInds       <- grep("externalLinks/externalLink[0-9]+.xml",       self$workbook.xml.rels)
      sharedStringsInd <- grep("sharedStrings.xml",                          self$workbook.xml.rels)
      tableInds        <- grep("table[0-9]+.xml",                            self$workbook.xml.rels)
      personInds       <- grep("person.xml",                                 self$workbook.xml.rels)


      ## Reordering of workbook.xml.rels
      ## don't want to re-assign rIds for pivot tables or slicer caches
      pivotNode        <- grep("pivotCache/pivotCacheDefinition[0-9].xml",  self$workbook.xml.rels, value = TRUE)
      slicerNode       <- grep("slicerCache[0-9]+.xml",                     self$workbook.xml.rels, value = TRUE)

      ## Reorder children of workbook.xml.rels
      self$workbook.xml.rels <-
        self$workbook.xml.rels[c(
          sheetInds,
          extRefInds,
          themeInd,
          connectionsInd,
          stylesInd,
          sharedStringsInd,
          tableInds,
          personInds
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
        self$workbook$bookViews <-
        sprintf(
          '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="13125" windowHeight="6105" firstSheet="%s" activeTab="%s"/></bookViews>',
          visible_sheet_index - 1L,
          visible_sheet_index - 1L
        )

      self$worksheets[[visible_sheet_index]]$sheetViews <-
        sub(
          '( tabSelected="0")|( tabSelected="false")',
          ' tabSelected="1"',
          self$worksheets[[visible_sheet_index]]$sheetViews,
          ignore.case = TRUE
        )
      if (nSheets > 1) {
        for (i in setdiff(seq_len(nSheets), visible_sheet_index)) {
          self$worksheets[[i]]$sheetViews <-
            sub(
              ' tabSelected="(1|true|false|0)"',
              ' tabSelected="0"',
              self$worksheets[[i]]$sheetViews,
              ignore.case = TRUE
            )
        }
      }


      if (length(self$workbook$definedNames)) {
        # TODO consider self$get_sheet_names() which orders the sheet names?
        sheets <- self$sheet_names[self$sheetOrder]

        belongTo <- get_named_regions(self)$sheets

        ## sheets is in re-ordered order (order it will be displayed)
        newId <- match(belongTo, sheets) - 1L
        oldId <- as.integer(reg_match0(self$workbook$definedNames, '(?<= localSheetId=")[0-9]+'))

        for (i in seq_along(self$workbook$definedNames)) {
          if (!is.na(newId[i])) {
            self$workbook$definedNames[[i]] <-
              gsub(
                sprintf('localSheetId=\"%s\"', oldId[i]),
                sprintf('localSheetId=\"%s\"', newId[i]),
                self$workbook$definedNames[[i]],
                fixed = TRUE
              )
          }
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


max_sheet_id <- function(wb) {
  if (!length(wb$workbook$sheets)) {
    return(1L)
  }

  max(get_sheet_id(wb), 0L, na.rm = TRUE) + 1L
}

get_sheet_id <- function(wb, index = NULL) {
  get_wb_sheet_id(wb, '(?<=sheetId=")[0-9]+', i = index)
}

get_r_id <- function(wb, index = NULL) {
  get_wb_sheet_id(wb, '(?<= r:id="rId)[0-9]+', i = index)
}

get_wb_sheet_id <- function(wb, pattern, i = NULL) {
  i <- i %||% seq_along(wb$workbook$sheets)
  id <- reg_match0(wb$workbook$sheets[i], pattern)
  as.integer(unlist(id))
}

# TODO Does this need to be checked?  No sheet name can be NA right?
# res <- self$sheet_names[ind]; stopifnot(!anyNA(ind))

#' Get sheet name
#'
#' @param wb a [wbWorkbook] object
#' @param index Sheet name index
#' @return The sheet index
#' @export
wb_get_sheet_name = function(wb, index = NULL) {
  index <- index %||% seq_along(wb$sheet_names)

  # index should be integer like
  stopifnot(is_integer_ish(index))

  n <- length(wb$sheet_names)

  if (any(index > n)) {
    stop("Invalid sheet index. Workbook ", n, " sheet(s)", call. = FALSE)
  }

  wb$sheet_names[index]
}
