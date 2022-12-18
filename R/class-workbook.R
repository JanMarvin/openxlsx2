

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

    #' @field is_chartsheet is_chartsheet
    is_chartsheet = logical(),

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
      self$apps <- character()
      self$charts <- list()
      self$is_chartsheet <- logical()

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
      # self$drawings_vml <- list()

      self$embeddings <- NULL
      self$externalLinks <- NULL
      self$externalLinksRels <- NULL

      self$headFoot <- NULL

      self$media <- list()
      self$metadata <- NULL

      self$persons <- NULL

      self$pivotTables <- NULL
      self$pivotTables.xml.rels <- NULL
      self$pivotDefinitions <- NULL
      self$pivotRecords <- NULL
      self$pivotDefinitionsRels <- NULL

      self$queryTables <- NULL

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
      private$current_sheet <- 0L
      invisible(self)
    },

    #' @description
    #' Append a field. This method is intended for internal use
    #' @param field A valid field name
    #' @param value A value for the field
    append = function(field, value) {
      self[[field]] <- c(self[[field]], value)
      invisible(self)
    },

    #' @description
    #' Append to `self$workbook$sheets` This method is intended for internal use
    #' @param value A value for `self$workbook$sheets`
    append_sheets = function(value) {
      self$workbook$sheets <- c(self$workbook$sheets, value)
      invisible(self)
    },

    #' @description validate sheet
    #' @param sheet A character sheet name or integer location
    #' @returns The integer position of the sheet
    validate_sheet = function(sheet) {

      # workbook has no sheets
      if (!length(self$sheet_names)) {
        return(NA_integer_)
      }

      # write_comment uses wb_validate and bails otherwise
      if (inherits(sheet, "openxlsx2_waiver")) {
        sheet <- private$get_sheet_index(sheet)
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
    #' Add a chart sheet to the workbook
    #' @param sheet sheet
    #' @param tabColour tabColour
    #' @param zoom zoom
    #' @param visible visible
    #' @return The `wbWorkbook` object, invisibly
    add_chartsheet = function(
      sheet     = next_sheet(),
      tabColour = NULL,
      zoom      = 100,
      visible   = c("true", "false", "hidden", "visible", "veryhidden")
    ) {
      visible <- tolower(as.character(visible))
      visible <- match.arg(visible)

      # set up so that a single error can be reported on fail
      fail <- FALSE
      msg <- NULL

      private$validate_new_sheet(sheet)

      if (is_waiver(sheet)) {
        if (sheet == "current_sheet") {
          stop("cannot add worksheet to current sheet")
        }

        # TODO openxlsx2.sheet.default_name is undocumented. should incorporate
        # a better check for this
        sheet <- paste0(
          getOption("openxlsx2.sheet.default_name", "Sheet "),
          length(self$sheet_names) + 1L
        )
      }

      sheet <- as.character(sheet)
      sheet_name <- replace_legal_chars(sheet)
      private$validate_new_sheet(sheet_name)


      newSheetIndex <- length(self$worksheets) + 1L
      private$set_current_sheet(newSheetIndex)
      sheetId <- private$get_sheet_id_max() # checks for self$worksheet length

      self$append_sheets(
        sprintf(
          '<sheet name="%s" sheetId="%s" state="%s" r:id="rId%s"/>',
          sheet_name,
          sheetId,
          visible,
          newSheetIndex
        )
      )

      if (!is.null(tabColour)) {
        tabColour <- validateColour(tabColour, "Invalid tabColour in add_worksheet.")
      }

      if (!is.numeric(zoom)) {
        fail <- TRUE
        msg <- c(msg, "zoom must be numeric")
      }

      # nocov start
      if (zoom < 10) {
        zoom <- 10
      } else if (zoom > 400) {
        zoom <- 400
      }
      #nocov end

      self$append("worksheets",
        wbChartSheet$new(
          tabColour   = tabColour
        )
      )

      self$worksheets[[newSheetIndex]]$set_sheetview(
        workbookViewId = 0,
        zoomScale      = zoom,
        tabSelected    = newSheetIndex == 1
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
      new_drawings_idx <- length(self$drawings) + 1
      self$drawings[[new_drawings_idx]]      <- ""
      self$drawings_rels[[new_drawings_idx]] <- ""

      self$worksheets_rels[[newSheetIndex]]  <- genBaseSheetRels(newSheetIndex)
      self$is_chartsheet[[newSheetIndex]]    <- TRUE
      self$vml_rels[[newSheetIndex]]         <- list()
      self$vml[[newSheetIndex]]              <- list()

      self$append("sheetOrder", newSheetIndex)
      private$set_single_sheet_name(newSheetIndex, sheet_name, sheet)

      invisible(self)
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
      visible <- tolower(as.character(visible))
      visible <- match.arg(visible)
      orientation <- match.arg(orientation, c("portrait", "landscape"))

      # set up so that a single error can be reported on fail
      fail <- FALSE
      msg <- NULL

      private$validate_new_sheet(sheet)

      if (is_waiver(sheet)) {
        if (sheet == "current_sheet") {
          stop("cannot add worksheet to current sheet")
        }

        # TODO openxlsx2.sheet.default_name is undocumented. should incorporate
        # a better check for this
        sheet <- paste0(
          getOption("openxlsx2.sheet.default_name", "Sheet "),
          length(self$sheet_names) + 1L
        )
      }

      sheet <- as.character(sheet)
      sheet_name <- replace_legal_chars(sheet)
      private$validate_new_sheet(sheet_name)

      if (!is.logical(gridLines) | length(gridLines) > 1) {
        fail <- TRUE
        msg <- c(msg, "gridLines must be a logical of length 1.")
      }

      if (!is.null(tabColour)) {
        tabColour <- validateColour(tabColour, "Invalid tabColour in add_worksheet.")
      }

      if (!is.numeric(zoom)) {
        fail <- TRUE
        msg <- c(msg, "zoom must be numeric")
      }

      # nocov start
      if (zoom < 10) {
        zoom <- 10
      } else if (zoom > 400) {
        zoom <- 400
      }
      #nocov end

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

      if (fail) {
        stop(msg, call. = FALSE)
      }

      newSheetIndex <- length(self$worksheets) + 1L
      private$set_current_sheet(newSheetIndex)
      sheetId <- private$get_sheet_id_max() # checks for self$worksheet length

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

      self$append_sheets(
        sprintf(
          '<sheet name="%s" sheetId="%s" state="%s" r:id="rId%s"/>',
          sheet_name,
          sheetId,
          visible,
          newSheetIndex
        )
      )

      ## append to worksheets list
      self$append("worksheets",
        wbWorksheet$new(
          tabColour   = tabColour,
          oddHeader   = oddHeader,
          oddFooter   = oddFooter,
          evenHeader  = evenHeader,
          evenFooter  = evenFooter,
          firstHeader = firstHeader,
          firstFooter = firstFooter,
          paperSize   = paperSize,
          orientation = orientation,
          hdpi        = hdpi,
          vdpi        = vdpi,
          printGridLines = gridLines
        )
      )

      # NULL or TRUE/FALSE
      rightToLeft <- getOption("openxlsx2.rightToLeft")

      # set preselected set for sheetview
      self$worksheets[[newSheetIndex]]$set_sheetview(
        workbookViewId    = 0,
        zoomScale         = zoom,
        showGridLines     = gridLines,
        showRowColHeaders = rowColHeaders,
        tabSelected       = newSheetIndex == 1,
        rightToLeft       = rightToLeft
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
      new_drawings_idx <- length(self$drawings) + 1
      self$drawings[[new_drawings_idx]]      <- ""
      self$drawings_rels[[new_drawings_idx]] <- ""

      self$worksheets_rels[[newSheetIndex]]  <- genBaseSheetRels(newSheetIndex)
      self$vml_rels[[newSheetIndex]]         <- list()
      self$vml[[newSheetIndex]]              <- list()
      self$is_chartsheet[[newSheetIndex]]    <- FALSE
      self$comments[[newSheetIndex]]         <- list()
      self$threadComments[[newSheetIndex]]   <- list()

      self$append("sheetOrder", as.integer(newSheetIndex))
      private$set_single_sheet_name(newSheetIndex, sheet_name, sheet)

      invisible(self)
    },

    # TODO should this be as simple as: wb$wb_add_worksheet(wb$worksheets[[1]]$clone()) ?

    #' @description
    #' Clone a workbooksheet
    #' @param old name of worksheet to clone
    #' @param new name of new worksheet to add
    clone_worksheet = function(old = current_sheet(), new = next_sheet()) {
      private$validate_new_sheet(new)
      old <- private$get_sheet_index(old)

      newSheetIndex <- length(self$worksheets) + 1L
      private$set_current_sheet(newSheetIndex)
      sheetId <- private$get_sheet_id_max() # checks for length of worksheets

      if (!all(self$charts$chartEx == "")) {
        warning(
          "The file you have loaded contains chart extensions. At the moment,",
          " cloning worksheets can damage the output."
        )
      }

      # not the best but a quick fix
      new_raw <- new
      new <- replace_legal_chars(new)

      ## copy visibility from cloned sheet!
      visible <- rbindlist(xml_attr(self$workbook$sheets[[old]], "sheet"))$state

      ##  Add sheet to workbook.xml
      self$append_sheets(
        xml_node_create(
          "sheet",
          xml_attributes = c(
          name = new,
          sheetId = sheetId,
          state = visible,
          `r:id` = paste0("rId", newSheetIndex)
          )
        )
      )

      ## append to worksheets list
      self$append("worksheets", self$worksheets[[old]]$clone(deep = TRUE))

      ## update content_tyes
      ## add a drawing.xml for the worksheet
      # FIXME only add what is needed. If no previous drawing is found, don't
      # add a new one
      self$append("Content_Types", c(
        if (self$is_chartsheet[old]) {
          sprintf('<Override PartName="/xl/chartsheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>', newSheetIndex)
        } else {
          sprintf('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', newSheetIndex)
        }
      ))

      ## Update xl/rels
      self$append(
        "workbook.xml.rels",
        if (self$is_chartsheet[old]) {
          sprintf('<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet%s.xml"/>', newSheetIndex)
        } else {
          sprintf('<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>', newSheetIndex)
        }
      )

      ## create sheet.rels to simplify id assignment
      self$worksheets_rels[[newSheetIndex]] <- self$worksheets_rels[[old]]

      old_drawing_sheet <- NULL

      if (length(self$worksheets_rels[[old]])) {
        relship <- rbindlist(xml_attr(self$worksheets_rels[[old]], "Relationship"))
        relship$typ <- basename(relship$Type)
        old_drawing_sheet  <- as.integer(gsub("\\D+", "", relship$Target[relship$typ == "drawing"]))
      }

      if (length(old_drawing_sheet)) {

        new_drawing_sheet <- length(self$drawings) + 1

        self$drawings_rels[[new_drawing_sheet]] <- self$drawings_rels[[old_drawing_sheet]]

        # give each chart its own filename (images can re-use the same file, but charts can't)
        self$drawings_rels[[new_drawing_sheet]] <-
          # TODO Can this be simplified?  There's a bit going on here
          vapply(
            self$drawings_rels[[new_drawing_sheet]],
            function(rl) {
              # is rl here a length of 1?
              stopifnot(length(rl) == 1L) # lets find out...  if this fails, just remove it
              chartfiles <- reg_match(rl, "(?<=charts/)chart[0-9]+\\.xml")

              for (cf in chartfiles) {
                chartid <- nrow(self$charts) + 1L
                newname <- stri_join("chart", chartid, ".xml")
                old_chart <- as.integer(gsub("\\D+", "", cf))
                self$charts <- rbind(self$charts, self$charts[old_chart, ])

                # Read the chartfile and adjust all formulas to point to the new
                # sheet name instead of the clone source

                chart <- self$charts$chart[chartid]
                self$charts$rels[chartid] <- gsub("?drawing[0-9]+.xml", paste0("drawing", chartid, ".xml"), self$charts$rels[chartid])

                guard_ws <- function(x) {
                  if (grepl(" ", x)) x <- shQuote(x, type = "sh")
                  x
                }

                old_sheet_name <- guard_ws(self$sheet_names[[old]])
                new_sheet_name <- guard_ws(new)

                ## we need to replace "'oldname'" as well as "oldname"
                chart <- gsub(
                  old_sheet_name,
                  new_sheet_name,
                  chart,
                  perl = TRUE
                )

                self$charts$chart[chartid] <- chart

                # two charts can not point to the same rels
                if (self$charts$rels[chartid] != "") {
                  self$charts$rels[chartid] <- gsub(
                    stri_join(old_chart, ".xml"),
                    stri_join(chartid, ".xml"),
                    self$charts$rels[chartid]
                  )
                }

                rl <- gsub(stri_join("(?<=charts/)", cf), newname, rl, perl = TRUE)
              }

              rl

            },
            NA_character_,
            USE.NAMES = FALSE
          )

        # otherwise an empty drawings relationship is written
        if (identical(self$drawings_rels[[new_drawing_sheet]], character()))
          self$drawings_rels[[new_drawing_sheet]] <- list()


        self$drawings[[new_drawing_sheet]]       <- self$drawings[[old_drawing_sheet]]
      }

      ## TODO Currently it is not possible to clone a sheet with a slicer in a
      #  safe way. It will always result in a broken xlsx file which is fixable
      #  but will not contain a slicer.

      # most likely needs to add slicerCache for each slicer with updated names

      ## SLICERS

      rid <- as.integer(sub("\\D+", "", get_relship_id(obj = self$worksheets_rels[[newSheetIndex]], "slicer")))
      if (length(rid)) {

        warning("Cloning slicers is not yet supported. It will not appear on the sheet.")
        self$worksheets_rels[[newSheetIndex]] <- relship_no(obj = self$worksheets_rels[[newSheetIndex]], x = "slicer")

        newid <- length(self$slicers) + 1

        cloned_slicers <- self$slicers[[old]]
        slicer_attr <- xml_attr(cloned_slicers, "slicers")

        # Replace name with name_n. This will prevent the slicer from loading,
        # but the xlsx file is not broken
        slicer_child <- xml_node(cloned_slicers, "slicers", "slicer")
        slicer_df <- rbindlist(xml_attr(slicer_child, "slicer"))[c("name", "cache", "caption", "rowHeight")]
        slicer_df$name <- paste0(slicer_df$name, "_n")
        slicer_child <- df_to_xml("slicer", slicer_df)

        self$slicers[[newid]] <- xml_node_create("slicers", slicer_child, slicer_attr[[1]])

        self$worksheets_rels[[newSheetIndex]] <- c(
          self$worksheets_rels[[newSheetIndex]],
          sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicer\" Target=\"../slicers/slicer%s.xml\"/>",
                  rid,
                  newid)
        )

        self$Content_Types <- c(
          self$Content_Types,
          sprintf("<Override PartName=\"/xl/slicers/slicer%s.xml\" ContentType=\"application/vnd.ms-excel.slicer+xml\"/>", newid)
        )

      }

      # The IDs in the drawings array are sheet-specific, so within the new
      # cloned sheet the same IDs can be used => no need to modify drawings
      self$vml_rels[[newSheetIndex]]       <- self$vml_rels[[old]]
      self$vml[[newSheetIndex]]            <- self$vml[[old]]
      self$is_chartsheet[[newSheetIndex]]  <- self$is_chartsheet[[old]]
      self$comments[[newSheetIndex]]       <- self$comments[[old]]
      self$threadComments[[newSheetIndex]] <- self$threadComments[[old]]

      self$append("sheetOrder", as.integer(newSheetIndex))
      self$append("sheet_names", new)
      private$set_single_sheet_name(pos = newSheetIndex, clean = new, raw = new_raw)


      ############################
      ## DRAWINGS

      # if we have drawings to clone, remove every table reference from Relationship

      rid <- as.integer(sub("\\D+", "", get_relship_id(obj = self$worksheets_rels[[newSheetIndex]], x = "drawing")))

      if (length(rid)) {

        self$worksheets_rels[[newSheetIndex]] <- relship_no(obj = self$worksheets_rels[[newSheetIndex]], x = "drawing")

        self$worksheets_rels[[newSheetIndex]] <- c(
          self$worksheets_rels[[newSheetIndex]],
          sprintf(
            '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing%s.xml"/>',
            rid,
            new_drawing_sheet
          )
        )

      }

      ############################
      ## TABLES
      ## ... are stored in the $tables list, with the name and sheet as attr
      ## and in the worksheets[]$tableParts list. We also need to adjust the
      ## worksheets_rels and set the content type for the new table

      # if we have tables to clone, remove every table referece from Relationship
      rid <- as.integer(sub("\\D+", "", get_relship_id(obj = self$worksheets_rels[[newSheetIndex]], x = "table")))

      if (length(rid)) {

        self$worksheets_rels[[newSheetIndex]] <- relship_no(obj = self$worksheets_rels[[newSheetIndex]], x = "table")

        # make this the new sheets object
        tbls <- self$tables[self$tables$tab_sheet == old, ]
        if (NROW(tbls)) {

          # newid and rid can be different. ids must be unique
          newid <- max(as.integer(rbindlist(xml_attr(self$tables$tab_xml, "table"))$id)) + seq_along(rid)

          # add _n to all table names found
          tbls$tab_name <- stri_join(tbls$tab_name, "_n")
          tbls$tab_sheet <- newSheetIndex
          # modify tab_xml with updated name, displayName and id
          tbls$tab_xml <- vapply(seq_len(nrow(tbls)), function(x) {
            xml_attr_mod(tbls$tab_xml[x],
                         xml_attributes = c(name = tbls$tab_name[x],
                                            displayName = tbls$tab_name[x],
                                            id = newid[x])
            )
          },
          NA_character_
          )

          # add new tables to old tables
          self$tables <- rbind(
            self$tables,
            tbls
          )

          self$worksheets[[newSheetIndex]]$tableParts                    <- sprintf('<tablePart r:id="rId%s"/>', rid)
          attr(self$worksheets[[newSheetIndex]]$tableParts, "tableName") <- tbls$tab_name

          ## hint: Content_Types will be created once the sheet is written. no need to add tables there

          # increase tables.xml.rels
          self$append("tables.xml.rels", rep("", nrow(tbls)))

          # add table.xml to worksheet relationship
          self$worksheets_rels[[newSheetIndex]] <- c(
            self$worksheets_rels[[newSheetIndex]],
            sprintf(
              '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',
              rid,
              newid
            )
          )
        }

      }

      # TODO: The following items are currently NOT copied/duplicated for the cloned sheet:
      #   - Comments ???
      #   - Slicers

      invisible(self)
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
        na.strings
      ) {

      if (missing(na.strings)) na.strings <- substitute()

      write_data(
        wb              = self,
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
      invisible(self)
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
        sheet       = current_sheet(),
        x,
        startCol    = 1,
        startRow    = 1,
        dims        = rowcol_to_dims(startRow, startCol),
        xy          = NULL,
        colNames    = TRUE,
        rowNames    = FALSE,
        tableStyle  = "TableStyleLight9",
        tableName   = NULL,
        withFilter  = TRUE,
        sep         = ", ",
        firstColumn = FALSE,
        lastColumn  = FALSE,
        bandedRows  = TRUE,
        bandedCols  = FALSE,
        applyCellStyle = TRUE,
        removeCellStyle = FALSE,
        na.strings
    ) {

      if (missing(na.strings)) na.strings <- substitute()

      write_datatable(
        wb          = self,
        sheet       = sheet,
        x           = x,
        dims        = dims,
        startCol    = startCol,
        startRow    = startRow,
        xy          = xy,
        colNames    = colNames,
        rowNames    = rowNames,
        tableStyle  = tableStyle,
        tableName   = tableName,
        withFilter  = withFilter,
        sep         = sep,
        firstColumn = firstColumn,
        lastColumn  = lastColumn,
        bandedRows  = bandedRows,
        bandedCols  = bandedCols,
        applyCellStyle = applyCellStyle,
        removeCellStyle = removeCellStyle,
        na.strings  = na.strings
      )
      invisible(self)
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
      write_formula(
        wb       = self,
        sheet    = sheet,
        x        = x,
        startCol = startCol,
        startRow = startRow,
        dims     = dims,
        array    = array,
        xy       = xy,
        applyCellStyle = applyCellStyle,
        removeCellStyle = removeCellStyle
      )
      invisible(self)
    },

    #' @description add style
    #' @param style style
    #' @param style_name style_name
    #' @returns The `wbWorkbook` object
    add_style = function(style = NULL, style_name = NULL) {

      assert_class(style, "character")

      if (is.null(style_name)) {
        style_name <- deparse(substitute(style))
      } else {
        assert_class(style_name, "character")
      }

      self$styles_mgr$add(style, style_name)

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

      success <- dir.create(path = tmpDir, recursive = FALSE)
      if (!success) { # nocov start
        stop(sprintf("Failed to create temporary directory '%s'", tmpDir))
      } # nocov end

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


      ## Content types has entries of all xml files in the workbook
      ct <- self$Content_Types


      ## will always have drawings
      xlworksheetsDir     <- dir_create(tmpDir, "xl", "worksheets")
      xlworksheetsRelsDir <- dir_create(tmpDir, "xl", "worksheets", "_rels")
      xldrawingsDir       <- dir_create(tmpDir, "xl", "drawings")
      xldrawingsRelsDir   <- dir_create(tmpDir, "xl", "drawings", "_rels")
      xlchartsDir         <- dir_create(tmpDir, "xl", "charts")
      xlchartsRelsDir     <- dir_create(tmpDir, "xl", "charts", "_rels")

      ## xl/comments.xml
      if (nComments > 0 | nVML > 0) {

        cmts <- rbindlist(xml_attr(unlist(self$worksheets_rels), "Relationship"))
        cmts$target <- basename(cmts$Target)
        cmts$typ <- basename(cmts$Type)
        cmts <- cmts[cmts$typ == "comments", ]
        cmts$id <- as.integer(gsub("\\D+", "", cmts$target))

        sel <- vapply(self$comments, function(x) length(x) > 0, NA)
        comments <- self$comments[sel]

        if (length(cmts$id) != length(comments))
          warning("comments length != comments ids")

        # TODO use seq_len() or seq_along()?
        for (i in seq_along(comments)) {
          fn <- sprintf("comments%s.xml", cmts$id[i])

          write_comment_xml(
            comment_list = comments[[i]],
            file_name = file.path(tmpDir, "xl", fn)
          )
        }

        private$writeDrawingVML(xldrawingsDir, xldrawingsRelsDir)
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

        slicer <- self$slicers[self$slicers != ""]
        for (i in seq_along(slicer)) {
          write_file(
            body = slicer[i],
            fl = file.path(slicersDir, sprintf("slicer%s.xml", i))
          )
        }

        for (i in seq_along(self$slicerCaches)) {
          write_file(
            body = self$slicerCaches[[i]],
            fl = file.path(slicerCachesDir, sprintf("slicerCache%s.xml", i))
          )
        }
      }


      ## Write content

      # if custom is present, we need 4 relationships, otherwise 3

      Ids <- c("rId3", "rId2", "rId1")
      Types <- c(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
      )
      Targets <- c("docProps/app.xml", "docProps/core.xml", "xl/workbook.xml")

      if (length(self$custom)) {
        Ids <- c(Ids, "rId4")
        Types <- c(
          Types,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        )
        Targets <- c(Targets, "docProps/custom.xml")
      }

      relship <- df_to_xml("Relationship",
        data.frame(Id = Ids, Type = Types, Target = Targets, stringsAsFactors = FALSE)
      )


      ## write .rels
      write_file(
        head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n',
        body = pxml(relship),
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


      ## write core.xml
      if (length(self$custom)) {
        write_file(
          head = "",
          body = pxml(self$custom),
          tail = "",
          fl = file.path(docPropsDir, "custom.xml")
        )
      }

      ## write workbook.xml.rels
      write_file(
        head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        body = pxml(self$workbook.xml.rels),
        tail = "</Relationships>",
        fl = file.path(xlrelsDir, "workbook.xml.rels")
      )

      ## write tables

      ## update tables in content types (some have been added, some removed, get the final state)
      default <- xml_node(ct, "Default")
      override <- rbindlist(xml_attr(ct, "Override"))
      override$typ <- gsub(".xml$", "", basename(override$PartName))
      override <- override[!grepl("table", override$typ), ]
      override$typ <- NULL

      # TODO remove length() check since we have seq_along()
      if (any(self$tables$tab_act == 1)) {

        # TODO get table Id from table entry
        table_ids <- function() {
          z <- 0
          if (!all(identical(unlist(self$worksheets_rels), character()))) {
            relship <- rbindlist(xml_attr(unlist(self$worksheets_rels), "Relationship"))
            relship$typ <- basename(relship$Type)
            relship$tid <- as.numeric(gsub("\\D+", "", relship$Target))

            z <- sort(relship$tid[relship$typ == "table"])
          }
          z
        }

        tab_ids <- table_ids()

        for (i in seq_along(tab_ids)) {

          # select only active tabs. in future there should only be active tabs
          tabs <- self$tables[self$tables$tab_act == 1, ]

          if (NROW(tabs)) {
            write_file(
              body = pxml(tabs$tab_xml[i]),
              fl = file.path(xlTablesDir, sprintf("table%s.xml", tab_ids[i]))
            )

            ## add entry to content_types as well
            override <- rbind(
              override,
              # new entry for table
              c("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                sprintf("/xl/tables/table%s.xml", tab_ids[i]))
            )

            if (self$tables.xml.rels[[i]] != "") {
              write_file(
                body = self$tables.xml.rels[[i]],
                fl = file.path(xlTablesRelsDir, sprintf("table%s.xml.rels", tab_ids[i]))
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

      ## connections
      if (length(self$ctrlProps)) {
        ctrlPropsDir <- dir_create(tmpDir, "xl", "ctrlProps")

        for (i in seq_along(self$ctrlProps)) {
          write_file(body = self$ctrlProps[i], fl = file.path(ctrlPropsDir, sprintf("ctrlProp%i.xml", i)))
        }
      }

      if (length(self$customXml)) {
        customXmlDir     <- dir_create(tmpDir, "customXml")
        customXmlRelsDir <- dir_create(tmpDir, "customXml", "_rels")
        for (fl in self$customXml[!grepl(".xml.rels$", self$customXml)]) {
          file.copy(fl, customXmlDir, overwrite = TRUE)
        }
        for (fl in self$customXml[grepl(".xml.rels$", self$customXml)]) {
          file.copy(fl, customXmlRelsDir, overwrite = TRUE)
        }
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
      ct <- private$writeSheetDataXML(
        ct,
        xldrawingsDir,
        xldrawingsRelsDir,
        xlchartsDir,
        xlchartsRelsDir,
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
          #body = stri_join(set_sst(attr(self$sharedStrings, "text")), collapse = "", sep = " "),
          body = stri_join(self$sharedStrings, collapse = "", sep = ""),
          tail = "</sst>",
          fl = file.path(xlDir, "sharedStrings.xml")
        )
      } else {
        ## Remove relationship to sharedStrings
        ct <- ct[!grepl("sharedStrings", ct)]
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

      ## do not write updated content types to self
      # self$Content_Types <- ct

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
      if (length(unlist(self$styles_mgr$styles))) {
        write_file(
          head = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr9="http://schemas.microsoft.com/office/spreadsheetml/2016/revision9" mc:Ignorable="x14ac x16r2 xr xr9">',
          body = pxml(styleXML),
          tail = "</styleSheet>",
          fl = file.path(xlDir, "styles.xml")
        )
      } else {
       write_file(
         head = '',
         body = '<styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
         tail = '',
         fl = file.path(xlDir, "styles.xml")
       )
      }

      if (length(self$calcChain)) {
        write_file(
          head = '',
          body = pxml(self$calcChain),
          tail = "",
          fl = file.path(xlDir, "calcChain.xml")
        )
      }

      # write metadata file. required if cm attribut is set.
      if (length(self$metadata)) {
        write_file(
          head = '',
          body = self$metadata,
          tail = '',
          fl = file.path(xlDir, "metadata.xml")
        )
      }

      ## write workbook.xml
      workbookXML <- self$workbook
      workbookXML$sheets <- stri_join("<sheets>", pxml(workbookXML$sheets), "</sheets>")

      if (length(workbookXML$definedNames)) {
        workbookXML$definedNames <- stri_join("<definedNames>", pxml(workbookXML$definedNames), "</definedNames>")
      }

      # openxml 2.8.1 expects the following order of xml nodes. While we create this per default, it is not
      # assured that the order of entries is still valid when we write the file. Functions can change the
      # workbook order, therefore we have to make sure that the expected order is written.
      # Othterwise spreadsheet software will complain.
      workbook_openxml281 <- c(
        "fileVersion", "fileSharing", "workbookPr", "alternateContent", "revisionPtr", "absPath",
        "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences",
        "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
        "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"
      )

      write_file(
        head = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">',
        body = pxml(workbookXML[workbook_openxml281]),
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
    #' @param interactive If `FALSE` will throw a warning and not open the path.
    #'   This can be manually set to `TRUE`, otherwise when `NA` (default) uses
    #'   the value returned from [base::interactive()]
    #' @return The `wbWorkbook`, invisibly
    open = function(interactive = NA) {
      xl_open(self, interactive = interactive)
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

      ## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
      last_table_id <- function() {
        z <- 0

        if (!all(unlist(self$worksheets_rels) == "")) {
          relship <- rbindlist(xml_attr(unlist(self$worksheets_rels), "Relationship"))
          # assign("relship", relship, globalenv())
          relship$typ <- basename(relship$Type)
          relship$tid <- as.numeric(gsub("\\D+", "", relship$Target))
          if (any(relship$typ == "table"))
            z <- max(relship$tid[relship$typ == "table"])
        }

        z
      }

      id <- as.character(last_table_id() + 1) # otherwise will start at 0 for table 1 length indicates the last known
      sheet <- wb_validate_sheet(self, sheet)
      # get the next highest rid
      rid <- 1
      if (!all(identical(self$worksheets_rels[[sheet]], character()))) {
        rid <- max(as.integer(sub("\\D+", "", rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))[["Id"]]))) + 1
      }

      if (is.null(self$tables)) {
        nms <- NULL
        tSheets <- NULL
        tNames <- NULL
        tActive <- NULL
      } else {
        nms <- self$tables$tab_ref
        tSheets <- self$tables$tab_sheet
        tNames <- self$tables$tab_name
        tActive <- self$tables$tab_act
      }


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

      tab_xml_new <- xml_node_create(
          xml_name = "table",
          xml_children = c(autofilter, tableColumns, tableStyleXML),
          xml_attributes = table_attrs
      )

      self$tables <- data.frame(
        tab_name = c(tNames, tableName),
        tab_sheet = c(tSheets, sheet),
        tab_ref = c(nms, ref),
        tab_xml = c(self$tables$tab_xml, tab_xml_new),
        tab_act = c(self$tables$tab_act, 1),
        stringsAsFactors = FALSE
      )

      self$worksheets[[sheet]]$tableParts <- c(
        self$worksheets[[sheet]]$tableParts,
          sprintf('<tablePart r:id="rId%s"/>', rid)
      )
      attr(self$worksheets[[sheet]]$tableParts, "tableName") <- c(
        tNames[tSheets == sheet & tActive == 1],
        tableName
      )

      ## create a table.xml.rels
      self$append("tables.xml.rels", "")

      ## update worksheets_rels
      self$worksheets_rels[[sheet]] <- c(
        self$worksheets_rels[[sheet]],
        sprintf(
          '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',
          rid,
          id
        )
      )

      invisible(self)
    },


    ### base font ----

    #' @description
    #' Get the base font
    #' @return A list of of the font
    get_base_font = function() {
      baseFont <- self$styles_mgr$styles$fonts[[1]]

      sz     <- unlist(xml_attr(baseFont, "font", "sz"))
      colour <- unlist(xml_attr(baseFont, "font", "color"))
      name   <- unlist(xml_attr(baseFont, "font", "name"))

      if (length(sz[[1]]) == 0) {
        sz <- list("val" = "11")
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
    set_base_font = function(fontSize = 11, fontColour = wb_colour(theme = "1"), fontName = "Calibri") {
      if (fontSize < 0) stop("Invalid fontSize")
      if (is.character(fontColour) && is.null(names(fontColour))) fontColour <- wb_colour(fontColour)
      self$styles_mgr$styles$fonts[[1]] <- create_font(sz = as.character(fontSize), color = fontColour, name = fontName)
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

      wbv <- self$workbook$bookViews

      if (is.null(wbv)) {
        wbv <- xml_node_create("workbookView")
      } else {
        wbv <- xml_node(wbv, "bookViews", "workbookView")
      }

      wbv <- xml_attr_mod(
        wbv,
        xml_attributes = c(
          activeTab              = as_xml_attr(activeTab),
          autoFilterDateGrouping = as_xml_attr(autoFilterDateGrouping),
          firstSheet             = as_xml_attr(firstSheet),
          minimized              = as_xml_attr(minimized),
          showHorizontalScroll   = as_xml_attr(showHorizontalScroll),
          showSheetTabs          = as_xml_attr(showSheetTabs),
          showVerticalScroll     = as_xml_attr(showVerticalScroll),
          tabRatio               = as_xml_attr(tabRatio),
          visibility             = as_xml_attr(visibility),
          windowHeight           = as_xml_attr(windowHeight),
          windowWidth            = as_xml_attr(windowWidth),
          xWindow                = as_xml_attr(xWindow),
          yWindow                = as_xml_attr(yWindow)
        )
      )

      self$workbook$bookViews <- xml_node_create(
        "bookViews",
        xml_children = wbv
      )

      invisible(self)
    },

    ### sheet names ----

    #' @description Get sheet names
    #' @returns A `named` `character` vector of sheet names in their order.  The
    #'   names represent the original value of the worksheet prior to any
    #'   character substitutions.
    get_sheet_names = function() {
      res <- self$sheet_names
      names(res) <- private$original_sheet_names
      res[self$sheetOrder]
    },

    #' @description
    #' Sets a sheet name
    #' @param old Old sheet name
    #' @param new New sheet name
    #' @return The `wbWorkbook` object, invisibly
    set_sheet_names = function(old = NULL, new) {
      # assume all names.  Default values makes the test check for wrappers a
      # little weird
      old <- old %||% seq_along(self$sheet_names)

      if (identical(old, new)) {
        return(self)
      }

      if (!length(self$worksheets)) {
        stop("workbook does not contain any sheets")
      }

      if (length(old) != length(new)) {
        stop("`old` and `new` must be the same length")
      }

      pos <- private$get_sheet_index(old)
      new_raw <- as.character(new)
      new_name <- replace_legal_chars(new_raw)

      if (identical(self$sheet_names[pos], new_name)) {
        return(self)
      }

      bad <- duplicated(tolower(new))
      if (any(bad)) {
        stop("Sheet names cannot have duplicates: ", toString(new[bad]))
      }

      # should be able to pull this out into a single private function
      for (i in seq_along(pos)) {
        private$validate_new_sheet(new_name[i])
        private$set_single_sheet_name(pos[i], new_name[i], new_raw[i])
        # TODO move this work into private$set_single_sheet_name()

        ## Rename in workbook
        sheetId <- private$get_sheet_id(type = "sheetId", pos[i])
        rId <- private$get_sheet_id(type = 'rId', pos[i])
        self$workbook$sheets[[pos[i]]] <-
          sprintf(
            '<sheet name="%s" sheetId="%s" r:id="rId%s"/>',
            new_name[i],
            sheetId,
            rId
          )

        ## rename defined names
        if (length(self$workbook$definedNames)) {
          ind <- get_named_regions(self)$sheets == old
          if (any(ind)) {
            nn <- sprintf("'%s'", new_name[i])
            nn <- stringi::stri_replace_all_fixed(self$workbook$definedName[ind], old, nn)
            nn <- stringi::stri_replace_all(nn, "'+", "'")
            self$workbook$definedNames[ind] <- nn
          }
        }
      }

      invisible(self)
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
    #' @param hidden hidden
    #' @return The `wbWorkbook` object, invisibly
    set_row_heights = function(sheet = current_sheet(), rows, heights = NULL, hidden = FALSE) {
      sheet <- private$get_sheet_index(sheet)

      # TODO move to wbWorksheet method

      # create all A columns so that row_attr is available.
      # Someone thought that it would be a splendid idea, if
      # all row_attr needs to match cc. This is fine, though
      # it brings the downside that these cells have to be
      # initialized.
      dims <- rowcol_to_dims(rows, 1)
      private$do_cell_init(sheet, dims)

      row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr
      sel <- match(rows, row_attr$r)

      if (!is.null(heights)) {
        if (length(rows) > length(heights)) {
          heights <- rep(heights, length.out = length(rows))
        }

        if (length(heights) > length(rows)) {
          stop("Greater number of height values than rows.")
        }

        row_attr[sel, "ht"] <- as.character(as.numeric(heights))
        row_attr[sel, "customHeight"] <- "1"
      }

      ## hide empty rows per default
      # xml_attr_mod(
      #   wb$worksheets[[1]]$sheetFormatPr,
      #   xml_attributes = c(zeroHeight = "1")
      # )

      if (hidden) {
        row_attr[sel, "hidden"] <- "1"
      }

      self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr

      invisible(self)
    },

    #' @description
    #' Sets a row height for a sheet
    #' @param sheet sheet
    #' @param rows rows
    #' @return The `wbWorkbook` object, invisibly
    remove_row_heights = function(sheet = current_sheet(), rows) {
      sheet <- private$get_sheet_index(sheet)

      row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

      if (is.null(row_attr)) {
        warning("There are no initialized rows on this sheet")
        return(invisible(self))
      }

      sel <- match(rows, row_attr$r)
      row_attr[sel, "ht"] <- ""
      row_attr[sel, "customHeight"] <- ""

      self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr

      invisible(self)
    },

    ## columns ----

    #' description
    #' creates column object for worksheet
    #' @param sheet sheet
    #' @param n n
    #' @param beg beg
    #' @param end end
    createCols = function(sheet = current_sheet(), n, beg, end) {
       sheet <- private$get_sheet_index(sheet)
       self$worksheets[[sheet]]$cols_attr <- df_to_xml("col", empty_cols_attr(n, beg, end))
    },

    #' @description
    #' Group cols
    #' @param sheet sheet
    #' @param cols cols
    #' @param collapsed collapsed
    #' @param levels levels
    #' @return The `wbWorkbook` object, invisibly
    group_cols = function(sheet = current_sheet(), cols, collapsed = FALSE, levels = NULL) {
      sheet <- private$get_sheet_index(sheet)

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
    ungroup_cols = function(sheet = current_sheet(), cols) {
      sheet <- private$get_sheet_index(sheet)

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

      invisible(self)
    },

    #' @description Remove row heights from a worksheet
    #' @param sheet A name or index of a worksheet
    #' @param cols Indices of columns to remove custom width (if any) from.
    #' @return The `wbWorkbook` object, invisibly
    remove_col_widths = function(sheet = current_sheet(), cols) {
      sheet <- private$get_sheet_index(sheet)

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
    set_col_widths = function(sheet = current_sheet(), cols, widths = 8.43, hidden = FALSE) {
      sheet <- private$get_sheet_index(sheet)

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
      col_width <- widths[ok]
      hidden <- hidden[ok]
      cols <- cols[ok]

      col_df <- self$worksheets[[sheet]]$unfold_cols()
      base_font <- wb_get_base_font(self)

      if (any(widths == "auto")) {
        df <- wb_to_df(self, sheet = sheet, cols = cols, colNames = FALSE)
        # TODO format(x) might not be the way it is formatted in the xlsx file.
        col_width <- vapply(df, function(x) max(nchar(format(x))), NA_real_)
      }


      # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column
      widths <- calc_col_width(base_font = base_font, col_width = col_width)

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
      invisible(self)
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
    group_rows = function(sheet = current_sheet(), rows, collapsed = FALSE, levels = NULL) {
      sheet <- private$get_sheet_index(sheet)

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
      sheet <- private$get_sheet_index(sheet)

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
    ungroup_rows = function(sheet = current_sheet(), rows) {
      sheet <- private$get_sheet_index(sheet)

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

      invisible(self)
    },

    #' @description
    #' Remove a worksheet
    #' @param sheet The worksheet to delete
    #' @return The `wbWorkbook` object, invisibly
    remove_worksheet = function(sheet = current_sheet()) {
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

      sheet       <- private$get_sheet_index(sheet)
      sheet_names <- self$sheet_names
      nSheets     <- length(sheet_names)
      sheet_names <- sheet_names[[sheet]]

      ## definedNames
      if (length(self$workbook$definedNames)) {
        # wb_validate_sheet() makes sheet an integer
        # so we need to remove this before getting rid of the sheet names
        self$workbook$definedNames <- self$workbook$definedNames[
          !get_nr_from_definedName(self)$sheets %in% self$sheet_names[sheet]
        ]
      }

      self$remove_named_region(sheet)
      self$sheet_names <- self$sheet_names[-sheet]
      private$original_sheet_names <- private$original_sheet_names[-sheet]

      # if a sheet has no relationships, we xml_rels will not contain data
      if (!is.null(self$worksheets_rels[[sheet]])) {

        xml_rels <- rbindlist(
          xml_attr(self$worksheets_rels[[sheet]], "Relationship")
        )

        if (nrow(xml_rels) && ncol(xml_rels)) {
          xml_rels$type   <- basename(xml_rels$Type)
          xml_rels$target <- basename(xml_rels$Target)
          xml_rels$target[xml_rels$type == "hyperlink"] <- ""
          xml_rels$target_ind <- as.numeric(gsub("\\D+", "", xml_rels$target))
        }

        comment_id    <- xml_rels$target_ind[xml_rels$type == "comments"]
        # TODO not every sheet has a drawing. this originates from a time where
        # every sheet created got a drawing assigned.
        drawing_id    <- xml_rels$target_ind[xml_rels$type == "drawing"]
        pivotTable_id <- xml_rels$target_ind[xml_rels$type == "pivotTable"]
        table_id      <- xml_rels$target_ind[xml_rels$type == "table"]
        thrComment_id <- xml_rels$target_ind[xml_rels$type == "threadedComment"]
        vmlDrawing_id <- xml_rels$target_ind[xml_rels$type == "vmlDrawing"]

        # NULL the sheets
        if (length(comment_id))    self$comments[[comment_id]]          <- NULL
        if (length(drawing_id))    self$drawings[[drawing_id]]          <- ""
        if (length(drawing_id))    self$drawings_rels[[drawing_id]]     <- ""
        if (length(thrComment_id)) self$threadComments[[thrComment_id]] <- NULL
        if (length(vmlDrawing_id)) self$vml[[vmlDrawing_id]]            <- NULL
        if (length(vmlDrawing_id)) self$vml_rels[[vmlDrawing_id]]       <- NULL

        #### Modify Content_Types
        ## remove last drawings(sheet).xml from Content_Types
        drawing_name <- xml_rels$target[xml_rels$type == "drawing"]
        if (!is.null(drawing_name) && !identical(drawing_name, character()))
          self$Content_Types <- grep(drawing_name, self$Content_Types, invert = TRUE, value = TRUE)

      }

      self$is_chartsheet <- self$is_chartsheet[-sheet]

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

      sel <- self$tables$tab_sheet == sheet
      # tableName is a character Vector with an attached name Vector.
      if (any(sel)) {
        self$tables$tab_name[sel] <- paste0(self$tables$tab_name[sel], "_openxlsx_deleted")
        tab_sheet <- self$tables$tab_sheet
        tab_sheet[sel] <- 0
        tab_sheet[tab_sheet > sheet] <- tab_sheet[tab_sheet > sheet] - 1L
        self$tables$tab_sheet <- tab_sheet
        self$tables$tab_ref[sel] <- ""
        self$tables$tab_xml[sel] <- ""

        # deactivate sheet
        self$tables$tab_act[sel] <- 0
      }

      # ## drawing will always be the first relationship
      # if (nSheets > 1) {
      #   for (i in seq_len(nSheets - 1L)) {
      #     # did this get updated from length of 3 to 2?
      #     #self$worksheets_rels[[i]][1:2] <- genBaseSheetRels(i)
      #     rel <- rbindlist(xml_attr(self$worksheets_rels[[i]], "Relationship"))
      #     if (nrow(rel) && ncol(rel)) {
      #       if (any(basename(rel$Type) == "drawing")) {
      #         rel$Target[basename(rel$Type) == "drawing"] <- sprintf("../drawings/drawing%s.xml", i)
      #       }
      #       if (is.null(rel$TargetMode)) rel$TargetMode <- ""
      #       self$worksheets_rels[[i]] <- df_to_xml("Relationship", rel[c("Id", "Type", "Target", "TargetMode")])
      #     }
      #   }
      # } else {
      #   self$worksheets_rels <- list()
      # }

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
    #' @param errorStyle The icon shown and the options how to deal with such inputs. Default "stop" (cancel), else "information" (prompt popup) or "warning" (prompt accept or change input)
    #' @param errorTitle The error title
    #' @param error The error text
    #' @param promptTitle The prompt title
    #' @param prompt The prompt text
    #' @returns The `wbWorkbook` object
    add_data_validation = function(
      sheet = current_sheet(),
      cols,
      rows,
      type,
      operator,
      value,
      allowBlank = TRUE,
      showInputMsg = TRUE,
      showErrorMsg = TRUE,
      errorStyle = NULL,
      errorTitle = NULL,
      error = NULL,
      promptTitle = NULL,
      prompt = NULL
    ) {

      sheet <- private$get_sheet_index(sheet)

      ## rows and cols
      if (!is.numeric(cols)) {
        cols <- col2int(cols)
      }
      rows <- as.integer(rows)

      assert_class(allowBlank, "logical")
      assert_class(showInputMsg, "logical")
      assert_class(showErrorMsg, "logical")

      ## check length of value
      if (length(value) > 2) {
        stop("value argument must be length <= 2")
      }

      valid_types <- c(
        "custom",
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

      if (!tolower(type) %in% c("custom", "list")) {
        if (!tolower(operator) %in% tolower(valid_operators)) {
          stop("Invalid 'operator' argument!")
        }

        operator <- valid_operators[tolower(valid_operators) %in% tolower(operator)][1]
      } else if (tolower(type) == "custom") {
        operator <- NULL
      } else {
        operator <- "between" ## ignored
      }

      ## All inputs validated

      type <- valid_types[tolower(valid_types) %in% tolower(type)][1]

      ## check input combinations
      if ((type == "date") && !inherits(value, "Date")) {
        stop("If type == 'date' value argument must be a Date vector.")
      }

      if ((type == "time") && !inherits(value, c("POSIXct", "POSIXt"))) {
        stop("If type == 'time' value argument must be a POSIXct or POSIXlt vector.")
      }


      value <- head(value, 2)
      allowBlank <- as.character(as.integer(allowBlank[1]))
      showInputMsg <- as.character(as.integer(showInputMsg[1]))
      showErrorMsg <- as.character(as.integer(showErrorMsg[1]))

      # prepare for worksheet
      origin <- get_date_origin(self, origin = TRUE)

      sqref <- stri_join(
        get_cell_refs(data.frame(
          "x" = c(min(rows), max(rows)),
          "y" = c(min(cols), max(cols))
        )),
        sep = " ",
        collapse = ":"
      )

      if (type == "list") {
        operator <- NULL
      }

      self$worksheets[[sheet]]$.__enclos_env__$private$data_validation(
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
        prompt       = prompt,
        origin       = origin,
        sqref        = sqref
      )

      invisible(self)
    },

    #' @description
    #' Set cell merging for a sheet
    #' @param sheet sheet
    #' @param rows,cols Row and column specifications.
    #' @return The `wbWorkbook` object, invisibly
    merge_cells = function(sheet = current_sheet(), rows = NULL, cols = NULL) {
      sheet <- private$get_sheet_index(sheet)

      # TODO send to wbWorksheet() method
      # self$worksheets[[sheet]]$merge_cells(rows = rows, cols = cols)
      # invisible(self)

      rows <- range(as.integer(rows))
      cols <- range(as.integer(cols))

      sqref <- paste0(int2col(cols), rows)
      sqref <- stri_join(sqref, collapse = ":", sep = " ")

      current <- rbindlist(xml_attr(xml = self$worksheets[[sheet]]$mergeCells, "mergeCell"))$ref

      # regmatch0 will return character(0) when x is NULL
      if (length(current)) {

        new_merge     <- unname(unlist(dims_to_dataframe(sqref, fill = TRUE)))
        current_cells <- lapply(current, function(x) unname(unlist(dims_to_dataframe(x, fill = TRUE))))
        intersects    <- vapply(current_cells, function(x) any(x %in% new_merge), NA)

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
      private$append_sheet_field(sheet, "mergeCells", sprintf('<mergeCell ref="%s"/>', sqref))
      invisible(self)
    },

    #' @description
    #' Removes cell merging for a sheet
    #' @param sheet sheet
    #' @param rows,cols Row and column specifications.
    #' @return The `wbWorkbook` object, invisibly
    unmerge_cells = function(sheet = current_sheet(), rows = NULL, cols = NULL) {
      sheet <- private$get_sheet_index(sheet)

      rows <- range(as.integer(rows))
      cols <- range(as.integer(cols))

      sqref <- paste0(int2col(cols), rows)
      sqref <- stri_join(sqref, collapse = ":", sep = " ")

      current <- rbindlist(xml_attr(xml = self$worksheets[[sheet]]$mergeCells, "mergeCell"))$ref

      if (!is.null(current)) {
        new_merge     <- unname(unlist(dims_to_dataframe(sqref, fill = TRUE)))
        current_cells <- lapply(current, function(x) unname(unlist(dims_to_dataframe(x, fill = TRUE))))
        intersects    <- vapply(current_cells, function(x) any(x %in% new_merge), NA)

        # Remove intersection
        self$worksheets[[sheet]]$mergeCells <- self$worksheets[[sheet]]$mergeCells[!intersects]
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
      sheet = current_sheet(),
      firstActiveRow = NULL,
      firstActiveCol = NULL,
      firstRow = FALSE,
      firstCol = FALSE
    ) {
      # TODO rename to setFreezePanes?

      # fine to do the validation before the actual check to prevent other errors
      sheet <- private$get_sheet_index(sheet)

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

    ## comment ----

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
        comment) {

      if (!missing(dims)) {
        xy <- unlist(dims_to_rowcol(dims))
        col <- xy[[1]]
        row <- as.integer(xy[[2]])
      }

      write_comment(
        wb = self,
        sheet = sheet,
        col = col,
        row = row,
        comment = comment
      ) # has no use: xy

      invisible(self)
    },

    #' @description Remove comment
    #' @param sheet sheet
    #' @param col column to apply the comment
    #' @param row row to apply the comment
    #' @param dims row and column as spreadsheet dimension, e.g. "A1"
    #' @param gridExpand Remove all comments inside the grid. Similar to dims "A1:B2"
    #' @returns The `wbWorkbook` object
    remove_comment = function(
      sheet = current_sheet(),
      col,
      row,
      dims  = rowcol_to_dims(row, col),
      gridExpand = TRUE
    ) {

      if (!missing(dims)) {
        xy <- unlist(dims_to_rowcol(dims))
        col <- xy[[1]]
        row <- as.integer(xy[[2]])
        # with gridExpand this is always true
        gridExpand <- TRUE
      }

      remove_comment(wb = self, sheet = sheet, col = col, row = row, gridExpand = TRUE)

      invisible(self)
    },

    ## conditional formatting ----

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
      sheet = current_sheet(),
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
      width  <- as.integer(round(width * 914400L, 0)) # (EMUs per inch)
      height <- as.integer(round(height * 914400L, 0)) # (EMUs per inch)

      sheet <- private$get_sheet_index(sheet)

      # TODO tools::file_ext() ...
      imageType <- regmatches(file, gregexpr("\\.[a-zA-Z]*$", file))
      imageType <- gsub("^\\.", "", imageType)
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

      # with userShape we might need to skip one ahead
      found <- private$get_drawingsref()
      if (sheet %in% found$sheet) {
        sheet_drawing <- found$id[found$sheet == sheet]
      } else {
        sel <- which.min(abs(found$sheet - sheet))
        sheet_drawing <- max(sheet, found$id[found$sheet == sel] + 1)
      }

      # add image to drawings_rels
      old_drawings_rels <- unlist(self$drawings_rels[[sheet_drawing]])
      if (all(old_drawings_rels == "")) old_drawings_rels <- NULL

      imageNo <- length(xml_node_name(self$drawings[[sheet_drawing]], "xdr:wsDr")) + 1L

      ## drawings rels (Reference from drawings.xml to image file in media folder)
      self$drawings_rels[[sheet_drawing]] <- c(
        old_drawings_rels,
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
        '<xdr:oneCellAnchor>',
        from,
        sprintf('<xdr:ext cx="%s" cy="%s"/>', width, height),
        genBasePic(imageNo),
        "<xdr:clientData/>",
        "</xdr:oneCellAnchor>"
      )

      xml_attr <- c(
        "xmlns:xdr" = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "xmlns:a" = "http://schemas.openxmlformats.org/drawingml/2006/main",
        "xmlns:r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      )

      drawing <- xml_node_create(
        "xdr:wsDr",
        xml_children = drawingsXML,
        xml_attributes = xml_attr
      )

      self$add_drawing(sheet, drawing)

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
      if (is.null(dev.list()[[1]])) {
        warning("No plot to insert.")
        return(invisible(self))
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

    #' @description Add xml drawing
    #' @param sheet sheet
    #' @param dims dims
    #' @param xml xml
    #' @returns The `wbWorkbook` object
    add_drawing = function(
      sheet = current_sheet(),
      xml,
      dims = NULL
    ) {
      sheet <- private$get_sheet_index(sheet)

      is_chartsheet <- self$is_chartsheet[sheet]

      xml <- read_xml(xml, pointer = FALSE)

      if (!(xml_node_name(xml) == "xdr:wsDr")) {
        error("xml needs to be a drawing.")
      }

      ext   <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:ext")
      grpSp <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:grpSp")
      grFrm <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:graphicFrame")
      clDt  <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:clientData")

      # include rvg graphic from specific position to one or two cell anchor
      if (!is.null(dims) && !is_chartsheet && xml_node_name(xml, "xdr:wsDr") == "xdr:absoluteAnchor") {

        twocell <- grepl(":", dims)

        if (twocell) {

          xdr_typ <- "xdr:twoCellAnchor"
          ext <- NULL

          dims_list <- strsplit(dims, ":")[[1]]
          cols <- col2int(dims_list)
          rows <- as.numeric(gsub("\\D+", "", dims_list))

          anchor <- paste0(
            "<xdr:from>",
            "<xdr:col>%s</xdr:col><xdr:colOff>0</xdr:colOff>",
            "<xdr:row>%s</xdr:row><xdr:rowOff>0</xdr:rowOff>",
            "</xdr:from>",
            "<xdr:to>",
            "<xdr:col>%s</xdr:col><xdr:colOff>0</xdr:colOff>",
            "<xdr:row>%s</xdr:row><xdr:rowOff>0</xdr:rowOff>",
            "</xdr:to>"
          )
          anchor <- sprintf(anchor, cols[1] - 1L, rows[1] - 1L, cols[2], rows[2])

        } else {

          xdr_typ <- "xdr:oneCellAnchor"

          cols <- col2int(dims)
          rows <- as.numeric(gsub("\\D+", "", dims))

          anchor <- paste0(
            "<xdr:from>",
            "<xdr:col>%s</xdr:col><xdr:colOff>0</xdr:colOff>",
            "<xdr:row>%s</xdr:row><xdr:rowOff>0</xdr:rowOff>",
            "</xdr:from>"
          )
          anchor <- sprintf(anchor, cols[1] - 1L, rows[1] - 1L)

        }

        xdr_typ_xml <- xml_node_create(
          xdr_typ,
          xml_children = c(
            anchor,
            ext,
            grpSp,
            grFrm,
            clDt
          )
        )

        xml <- xml_node_create(
          "xdr:wsDr",
          xml_attributes = c(
            "xmlns:a"   = "http://schemas.openxmlformats.org/drawingml/2006/main",
            "xmlns:r"   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "xmlns:pic" = "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "xmlns:xdr" = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          ),
          xml_children = xdr_typ_xml
        )
      }

      # usually sheet_drawing is sheet. If we have userShapes, sheet_drawing
      # can skip ahead. see test: unemployment-nrw202208.xlsx
      found <- private$get_drawingsref()
      if (sheet %in% found$sheet) {
        sheet_drawing <- found$id[found$sheet == sheet]
      } else {
        sel <- which.min(abs(found$sheet - sheet))
        sheet_drawing <- max(sheet, found$id[found$sheet == sel] + 1)
      }

      # check if sheet already contains drawing. if yes, try to integrate
      # our drawing into this else we only use our drawing.
      drawings <- self$drawings[[sheet_drawing]]
      if (drawings == "") {
        drawings <- xml
      } else {
        drawing_type <- xml_node_name(xml, "xdr:wsDr")
        xml_drawing <- xml_node(xml, "xdr:wsDr", drawing_type)
        drawings <- xml_add_child(drawings, xml_drawing)
      }
      self$drawings[[sheet_drawing]] <- drawings

      # get the correct next free relship id
      if (length(self$worksheets_rels[[sheet]]) == 0) {
        next_relship <- 1
        has_no_drawing <- TRUE
      } else {
        relship <- rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))
        relship$typ <- basename(relship$Type)
        next_relship <- as.integer(gsub("\\D+", "", relship$Id)) + 1L
        has_no_drawing <- !any(relship$typ == "drawing")
      }

      # if a drawing exisits, we already added ourself to it. Otherwise we
      # create a new drawing.
      if (has_no_drawing) {
        self$worksheets_rels[[sheet]] <- sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing%s.xml\"/>", next_relship, sheet_drawing)
        self$worksheets[[sheet]]$drawing <- sprintf("<drawing r:id=\"rId%s\"/>", next_relship)
      }

      invisible(self)
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
      dims = NULL
    ) {

      sheet <- private$get_sheet_index(sheet)
      is_chartsheet <- self$is_chartsheet[sheet]


      found <- private$get_drawingsref()
      if (sheet %in% found$sheet) {
        sheet_drawing <- found$id[found$sheet == sheet]
      } else {
        sel <- which.min(abs(found$sheet - sheet))
        sheet_drawing <- max(sheet, found$id[found$sheet == sel] + 1)
      }

      # chartsheets can not have multiple drawings
      if (is_chartsheet) {
        self$drawings[[sheet_drawing]]      <- ""
        self$drawings_rels[[sheet_drawing]] <- ""
      }

      next_chart <- NROW(self$charts) + 1

      chart <- data.frame(
        chart = xml,
        colors = colors1_xml,
        style = styleplot_xml,
        rels = chart1_rels_xml(next_chart),
        chartEx = "",
        relsEx = ""
      )

      self$charts <- rbind(self$charts, chart)

      len_drawing <- length(xml_node_name(self$drawings[[sheet_drawing]], "xdr:wsDr")) + 1L

      # create drawing. add it to self$drawings, the worksheet and rels
      self$add_drawing(
        sheet = sheet,
        xml = drawings(len_drawing),
        dims = dims
      )

      self$drawings_rels[[sheet]] <- drawings_rels(self$drawings_rels[[sheet]], next_chart)

      invisible(self)
    },

    #' @description Add mschart chart to the workbook
    #' @param sheet the sheet on which the graph will appear
    #' @param dims the dimensions where the sheet will appear
    #' @param graph mschart graph
    #' @returns The `wbWorkbook` object
    add_mschart = function(
      sheet = current_sheet(),
      dims = NULL,
      graph
    ) {

      requireNamespace("mschart")
      assert_class(graph, "ms_chart")

      sheetname <- private$get_sheet_name(sheet)

      # format.ms_chart is exported in mschart >= 0.4
      out_xml <- read_xml(
        format(
          graph,
          sheetname = sheetname,
          id_x = "64451212",
          id_y = "64453248"
        ),
        pointer = FALSE
      )

      # write the chart data to the workbook
      if (inherits(graph$data_series, "wb_data")) {
        self$
          add_chart_xml(sheet = sheet, xml = out_xml, dims = dims)
      } else {
        self$
          add_data(sheet = sheet, x = graph$data_series)$
          add_chart_xml(sheet = sheet, xml = out_xml, dims = dims)
      }
    },

    #' @description
    #' Prints the `wbWorkbook` object
    #' @return The `wbWorkbook` object, invisibly; called for its side-effects
    print = function() {
      exSheets <- self$get_sheet_names()
      nSheets <- length(exSheets)
      nImages <- length(self$media)
      nCharts <- length(self$charts)

      showText <- "A Workbook object.\n"

      ## worksheets
      if (nSheets > 0) {
        showText <- c(showText, "\nWorksheets:\n")
        sheetTxt <- sprintf("Sheets: %s", paste(names(exSheets), collapse = " "))

        showText <- c(showText, sheetTxt, "\n")
      } else {
        showText <-
          c(showText, "\nWorksheets:\n", "No worksheets attached\n")
      }

      if (nSheets > 0) {
        showText <-
          c(showText, sprintf(
            "Write order: %s",
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
      protect             = TRUE,
      password            = NULL,
      lockStructure       = FALSE,
      lockWindows         = FALSE,
      type                = c("1", "2", "4", "8"),
      fileSharing         = FALSE,
      username            = unname(Sys.info()["user"]),
      readOnlyRecommended = FALSE
    ) {

      if (!protect) {
        self$workbook$workbookProtection <- NULL
        return(self)
      }

      # match.arg() doesn't handle numbers too well
      type <- if (!is.character(type)) as.character(type)
      password <- if (is.null(password)) "" else hashPassword(password)

      # TODO: Shall we parse the existing protection settings and preserve all
      # unchanged attributes?

      if (fileSharing) {
        self$workbook$fileSharing <- xml_node_create(
          "fileSharing",
          xml_attributes = c(
            userName = username,
            readOnlyRecommended = if (readOnlyRecommended | type == "2") "1",
            reservationPassword = password
          )
        )
      }

      self$workbook$workbookProtection <- xml_node_create(
        "workbookProtection",
        xml_attributes = c(
          hashPassword = password,
          lockStructure = toString(as.numeric(lockStructure)),
          lockWindows = toString(as.numeric(lockWindows))
        )
      )

      self$workbook$apps <- xml_node_create("DocSecurity", type)
      invisible(self)
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

      sheet <- wb_validate_sheet(self, sheet)

      if (!protect) {
        # initializes as character()
        self$worksheets[[sheet]]$sheetProtection <- character()
        return(self)
      }

      all_props <- worksheet_lock_properties()

      if (!is.null(properties)) {
        # ensure only valid properties are listed
        properties <- match.arg(properties, all_props, several.ok = TRUE)
      }

      properties <- as.character(as.numeric(all_props %in% properties))
      names(properties) <- all_props

      if (!is.null(password))
        properties <- c(properties, password = hashPassword(password))

      self$worksheets[[sheet]]$sheetProtection <- xml_node_create(
        "sheetProtection",
        xml_attributes = c(
          sheet = "1",
          properties[properties != "0"]
        )
      )

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
      sheet <- private$get_sheet_index(sheet)
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
        '<pageSetup paperSize="%s" orientation="%s" scale = "%s" fitToWidth="%s" fitToHeight="%s" horizontalDpi="%s" verticalDpi="%s"/>',
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
          sheet = self$get_sheet_names()[[sheet]],
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
          sheet = self$get_sheet_names()[[sheet]],
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
        sheet <- self$get_sheet_names()[[sheet]]

        self$workbook$definedNames <- c(
          self$workbook$definedNames,
          sprintf('<definedName name="_xlnm.Print_Titles" localSheetId="%s">\'%s\'!%s,\'%s\'!%s</definedName>', localSheetId, sheet, cols, sheet, rows)
        )

      }

      invisible(self)
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
      sheet = current_sheet(),
      header      = NULL,
      footer      = NULL,
      evenHeader  = NULL,
      evenFooter  = NULL,
      firstHeader = NULL,
      firstFooter = NULL
    ) {
      sheet <- private$get_sheet_index(sheet)

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
      invisible(self)
    },

    #' @description get tables
    #' @param sheet sheet
    #' @returns The sheet tables.  `character()` if empty
    get_tables = function(sheet = current_sheet()) {
      if (length(sheet) != 1) {
        stop("sheet argument must be length 1")
      }

      if (is.null(self$tables)) {
        return(character())
      }

      sheet <- private$get_sheet_index(sheet)
      if (is.na(sheet)) stop("No such sheet in workbook")

      sel <- self$tables$tab_sheet == sheet & self$tables$tab_act == 1
      tables <- self$tables$tab_name[sel]
      refs <- self$tables$tab_ref[sel]

      if (length(tables)) {
        attr(tables, "refs") <- refs
      }

      return(tables)
    },


    #' @description remove tables
    #' @param sheet sheet
    #' @param table table
    #' @returns The `wbWorkbook` object
    remove_tables = function(sheet = current_sheet(), table) {
      if (length(table) != 1) {
        stop("table argument must be length 1")
      }

      ## delete table object and all data in it
      sheet <- private$get_sheet_index(sheet)

      if (!table %in% self$tables$tab_name) {
        stop(sprintf("table '%s' does not exist.", table), call. = FALSE)
      }

      ## delete table object (by flagging as deleted)
      inds <- self$tables$tab_sheet %in% sheet & self$tables$tab_name %in% table
      table_name_original <- self$tables$tab_name[inds]
      refs <- self$tables$tab_ref[inds]

      self$tables$tab_name[inds] <- paste0(table_name_original, "_openxlsx_deleted")
      self$tables$tab_ref[inds] <- ""
      self$tables$tab_sheet[inds] <- 0
      self$tables$tab_xml[inds] <- ""
      self$tables$tab_act[inds] <- 0

      ## delete reference from worksheet to table
      worksheet_table_names <- attr(self$worksheets[[sheet]]$tableParts, "tableName")
      to_remove <- which(worksheet_table_names == table_name_original)

      # (1) remove the rId from worksheet_rels
      rm_tab_rId <- rbindlist(xml_attr(self$worksheets[[sheet]]$tableParts[to_remove], "tablePart"))["r:id"]
      ws_rels <- self$worksheets_rels[[sheet]]
      is_rm_table <- grepl(rm_tab_rId, ws_rels)
      self$worksheets_rels[[sheet]] <- ws_rels[!is_rm_table]

      # (2) remove the rId from tableParts
      self$worksheets[[sheet]]$tableParts <- self$worksheets[[sheet]]$tableParts[-to_remove]
      attr(self$worksheets[[sheet]]$tableParts, "tableName") <- worksheet_table_names[-to_remove]


      ## Now delete data from the worksheet
      refs <- strsplit(refs, split = ":")[[1]]
      rows <- as.integer(gsub("[A-Z]", "", refs))
      rows <- seq(from = rows[1], to = rows[2], by = 1)

      cols <- col2int(refs)
      cols <- seq(from = cols[1], to = cols[2], by = 1)

      ## now delete data
      delete_data(wb = self, sheet = sheet, rows = rows, cols = cols)
      invisible(self)
    },

    #' @description add filters
    #' @param sheet sheet
    #' @param rows rows
    #' @param cols cols
    #' @returns The `wbWorkbook` object
    add_filter = function(sheet = current_sheet(), rows, cols) {
      sheet <- private$get_sheet_index(sheet)

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

      invisible(self)
    },

    #' @description remove filters
    #' @param sheet sheet
    #' @returns The `wbWorkbook` object
    remove_filter = function(sheet = current_sheet()) {
      for (s in private$get_sheet_index(sheet)) {
        self$worksheets[[s]]$autoFilter <- character()
      }

      invisible(self)
    },

    #' @description grid lines
    #' @param sheet sheet
    #' @param show show
    #' @param print print
    #' @returns The `wbWorkbook` object
    grid_lines = function(sheet = current_sheet(), show = FALSE, print = show) {
      sheet <- private$get_sheet_index(sheet)

      assert_class(show, "logical")
      assert_class(print, "logical")

      ## show
      sv <- self$worksheets[[sheet]]$sheetViews
      sv <- xml_attr_mod(sv, c(showGridLines = as_xml_attr(show)))
      self$worksheets[[sheet]]$sheetViews <- sv

      ## print
      if (print)
        self$worksheets[[sheet]]$set_print_options(gridLines = print, gridLinesSet = print)

      invisible(self)
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
      sheet <- private$get_sheet_index(sheet)

      if (!is.numeric(rows)) {
        stop("rows argument must be a numeric/integer vector")
      }

      if (!is.numeric(cols)) {
        stop("cols argument must be a numeric/integer vector")
      }

      localSheetId <- ""
      if (localSheet) localSheetId <- as.character(sheet)

      ## check name doesn't already exist
      ## named region

      definedNames <- rbindlist(xml_attr(self$workbook$definedNames, level1 = "definedName"))
      sel1 <- tolower(definedNames$name) == tolower(name)
      sel2 <- definedNames$localSheetId == localSheetId
      if (!is.null(definedNames$localSheetId)) {
        sel <- sel1 & sel2
      } else {
         sel <- sel1
      }
      match_dn <- which(sel)

      if (any(match_dn)) {
        if (overwrite)
          self$workbook$definedNames <- self$workbook$definedNames[-match_dn]
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

      if (localSheetId == "") localSheetId <- NULL

      private$create_named_region(
        ref1               = ref1,
        ref2               = ref2,
        name               = name,
        sheet              = self$sheet_names[sheet],
        localSheetId       = localSheetId,
        comment            = comment,
        customMenu         = customMenu,
        description        = description,
        is_function        = is_function,
        functionGroupId    = functionGroupId,
        help               = help,
        hidden             = hidden,
        localName          = localName,
        publishToServer    = publishToServer,
        statusBar          = statusBar,
        vbProcedure        = vbProcedure,
        workbookParameter  = workbookParameter,
        xml                = xml
      )

      invisible(self)
    },

    #' @description remove a named region
    #' @param sheet sheet
    #' @param name name
    #' @returns The `wbWorkbook` object
    remove_named_region = function(sheet = current_sheet(), name = NULL) {
      # get all nown defined names
      dn <- get_named_regions(self)

      if (is.null(name) && !is.null(sheet)) {
        sheet <- private$get_sheet_index(sheet)
        del <- dn$id[dn$sheet == sheet]
      } else if (!is.null(name) && is.null(sheet)) {
        del <- dn$id[dn$name == name]
      } else {
        sheet <- private$get_sheet_index(sheet)
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

      invisible(self)
    },

    #' @description set worksheet order
    #' @param sheets sheets
    #' @return The `wbWorkbook` object
    set_order = function(sheets) {
      sheets <- private$get_sheet_index(sheet = sheets)

      if (anyDuplicated(sheets)) {
        stop("`sheets` cannot have duplicates")
      }

      if (length(sheets) != length(self$worksheets)) {
        stop(sprintf("Worksheet order must be same length as number of worksheets [%s]", length(self$worksheets)))
      }

      if (any(sheets > length(self$worksheets))) {
        stop("Elements of order are greater than the number of worksheets")
      }

      self$sheetOrder <- sheets
      invisible(self)
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
    set_sheet_visibility = function(sheet = current_sheet(), value) {
      if (length(value) != length(sheet)) {
        stop("`value` and `sheet` must be the same length")
      }

      sheet <- private$get_sheet_index(sheet)

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
        return(invisible(self))
      }

      for (i in seq_along(self$worksheets)) {
        self$workbook$sheets[sheet[i]] <- gsub(exState0[i], value[i], self$workbook$sheets[sheet[i]], fixed = TRUE)
      }

      if (!any(self$get_sheet_visibility() %in% c("true", "visible"))) {
        warning("A workbook must have atleast 1 visible worksheet.  Setting first for visible")
        self$set_sheet_visibility(1, TRUE)
      }

      invisible(self)
    },

    ## page breaks ----

    #' @description Add a page break
    #' @param sheet sheet
    #' @param row row
    #' @param col col
    #' @returns The `wbWorkbook` object
    add_page_break = function(sheet = current_sheet(), row = NULL, col = NULL) {
      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$add_page_break(row = row, col = col)
      invisible(self)
    },

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
      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$clean_sheet(numbers = numbers, characters = characters,
        styles       = styles,
        merged_cells = merged_cells
      )
      invisible(self)
    },

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

      # TODO merge styles and if a style is already present, only add the newly
      # created border style

      # cc <- wb$worksheets[[sheet]]$sheet_data$cc
      # df_s <- as.data.frame(lapply(df, function(x) cc$c_s[cc$r %in% x]))

      df <- dims_to_dataframe(dims, fill = TRUE)
      sheet <- private$get_sheet_index(sheet)

      private$do_cell_init(sheet, dims)

      ### beg border creation
      full_single <- create_border(
        top = top_border, top_color = top_color,
        bottom = bottom_border, bottom_color = bottom_color,
        left = left_border, left_color = left_color,
        right = left_border, right_color = right_color
      )

      top_single <- create_border(
        top = top_border, top_color = top_color,
        bottom = inner_hgrid, bottom_color = inner_hcolor,
        left = left_border, left_color = left_color,
        right = left_border, right_color = right_color
      )

      middle_single <- create_border(
        top = inner_hgrid, top_color = inner_hcolor,
        bottom = inner_hgrid, bottom_color = inner_hcolor,
        left = left_border, left_color = left_color,
        right = left_border, right_color = right_color
      )

      bottom_single <- create_border(
        top = inner_hgrid, top_color = inner_hcolor,
        bottom = bottom_border, bottom_color = bottom_color,
        left = left_border, left_color = left_color,
        right = left_border, right_color = right_color
      )

      left_single <- create_border(
        top = top_border, top_color = top_color,
        bottom = bottom_border, bottom_color = bottom_color,
        left = left_border, left_color = left_color,
        right = inner_vgrid, right_color = inner_vcolor
      )

      right_single <- create_border(
        top = top_border, top_color = top_color,
        bottom = bottom_border, bottom_color = bottom_color,
        left = inner_vgrid, left_color = inner_vcolor,
        right = right_border, right_color = right_color
      )

      center_single <- create_border(
        top = top_border, top_color = top_color,
        bottom = bottom_border, bottom_color = bottom_color,
        left = inner_vgrid, left_color = inner_vcolor,
        right = inner_vgrid, right_color = inner_vcolor
      )

      top_left <- create_border(
        top = top_border, top_color = top_color,
        bottom = inner_hgrid, bottom_color = inner_hcolor,
        left = left_border, left_color = left_color,
        right = inner_vgrid, right_color = inner_vcolor
      )

      top_right <- create_border(
        top = top_border, top_color = top_color,
        bottom = inner_hgrid, bottom_color = inner_hcolor,
        left = inner_vgrid, left_color = inner_vcolor,
        right = left_border, right_color = right_color
      )

      bottom_left <- create_border(
        top = inner_hgrid, top_color = inner_hcolor,
        bottom = bottom_border, bottom_color = bottom_color,
        left = left_border, left_color = left_color,
        right = inner_vgrid, right_color = inner_vcolor
      )

      bottom_right <- create_border(
        top = inner_hgrid, top_color = inner_hcolor,
        bottom = bottom_border, bottom_color = bottom_color,
        left = inner_vgrid, left_color = inner_vcolor,
        right = left_border, right_color = right_color
      )

      top_center <- create_border(
        top = top_border, top_color = top_color,
        bottom = inner_hgrid, bottom_color = inner_hcolor,
        left = inner_vgrid, left_color = inner_vcolor,
        right = inner_vgrid, right_color = inner_vcolor
      )

      bottom_center <- create_border(
        top = inner_hgrid, top_color = inner_hcolor,
        bottom = bottom_border, bottom_color = bottom_color,
        left = inner_vgrid, left_color = inner_vcolor,
        right = inner_vgrid, right_color = inner_vcolor
      )

      middle_left <- create_border(
        top = inner_hgrid, top_color = inner_hcolor,
        bottom = inner_hgrid, bottom_color = inner_hcolor,
        left = left_border, left_color = left_color,
        right = inner_vgrid, right_color = inner_vcolor
      )

      middle_right <- create_border(
        top = inner_hgrid, top_color = inner_hcolor,
        bottom = inner_hgrid, bottom_color = inner_hcolor,
        left = inner_vgrid, left_color = inner_vcolor,
        right = right_border, right_color = right_color
      )

      inner_cell <- create_border(
        top = inner_hgrid, top_color = inner_hcolor,
        bottom = inner_hgrid, bottom_color = inner_hcolor,
        left = inner_vgrid, left_color = inner_vcolor,
        right = inner_vgrid, right_color = inner_vcolor
      )
      ### end border creation

      #
      # /* top_single    */
      # /* middle_single */
      # /* bottom_single */
      #

      # /* left_single --- center_single --- right_single */

      #
      # /* top_left   --- top_center   ---  top_right */
      # /*  -                                    -    */
      # /*  -                                    -    */
      # /*  -                                    -    */
      # /* left_middle                   right_middle */
      # /*  -                                    -    */
      # /*  -                                    -    */
      # /*  -                                    -    */
      # /* left_bottom - bottom_center - bottom_right */
      #

      ## beg cell references
      if (ncol(df) == 1 && nrow(df) == 1)
        dim_full_single <- df[1, 1]

      if (ncol(df) == 1 && nrow(df) >= 2) {
        dim_top_single <- df[1, 1]
        dim_bottom_single <- df[nrow(df), 1]
        if (nrow(df) >= 3) {
          mid <- df[, 1]
          dim_middle_single <- mid[!mid %in% c(dim_top_single, dim_bottom_single)]
        }
      }

      if (ncol(df) >= 2 && nrow(df) == 1) {
        dim_left_single <- df[1, 1]
        dim_right_single <- df[1, ncol(df)]
        if (ncol(df) >= 3) {
          ctr <- df[1, ]
          dim_center_single <- ctr[!ctr %in% c(dim_left_single, dim_right_single)]
        }
      }

      if (ncol(df) >= 2 && nrow(df) >= 2) {
        dim_top_left     <- df[1, 1]
        dim_bottom_left  <- df[nrow(df), 1]
        dim_top_right    <- df[1, ncol(df)]
        dim_bottom_right <- df[nrow(df), ncol(df)]

        if (nrow(df) >= 3) {
          top_mid <- df[, 1]
          bottom_mid <- df[, ncol(df)]

          dim_middle_left <- top_mid[!top_mid %in% c(dim_top_left, dim_bottom_left)]
          dim_middle_right <- bottom_mid[!bottom_mid %in% c(dim_top_right, dim_bottom_right)]
        }

        if (ncol(df) >= 3) {
          top_ctr <- df[1, ]
          bottom_ctr <- df[nrow(df), ]

          dim_top_center <- top_ctr[!top_ctr %in% c(dim_top_left, dim_top_right)]
          dim_bottom_center <- bottom_ctr[!bottom_ctr %in% c(dim_bottom_left, dim_bottom_right)]
        }

        if (ncol(df) > 2 && nrow(df) > 2) {
          t_row <- 1
          b_row <- nrow(df)
          l_row <- 1
          r_row <- ncol(df)
          dim_inner_cell <- as.character(unlist(df[c(-t_row, -b_row), c(-l_row, -r_row)]))
        }
      }
      ### end cell references

      # add some random string to the name. if called multiple times, new
      # styles will be created. We do not look for identical styles, therefor
      # we might create duplicates, but if a single style changes, the rest of
      # the workbook remains valid.
      smp <- random_string()
      s <- function(x) paste0(smp, "s", deparse(substitute(x)), seq_along(x))
      sfull_single <- paste0(smp, "full_single")
      stop_single <- paste0(smp, "full_single")
      sbottom_single <- paste0(smp, "bottom_single")
      smiddle_single <- paste0(smp, "middle_single")
      sleft_single <- paste0(smp, "left_single")
      sright_single <- paste0(smp, "right_single")
      scenter_single <- paste0(smp, "center_single")
      stop_left <- paste0(smp, "top_left")
      stop_right <- paste0(smp, "top_right")
      sbottom_left <- paste0(smp, "bottom_left")
      sbottom_right <- paste0(smp, "bottom_right")
      smiddle_left <- paste0(smp, "middle_left")
      smiddle_right <- paste0(smp, "middle_right")
      stop_center <- paste0(smp, "top_center")
      sbottom_center <- paste0(smp, "bottom_center")
      sinner_cell <- paste0(smp, "inner_cell")

      # ncol == 1
      if (ncol(df) == 1) {

        # single cell
        if (nrow(df) == 1) {
          self$styles_mgr$add(full_single, sfull_single)
          xf_prev <- get_cell_styles(self, sheet, dims)
          xf_full_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sfull_single))
          self$styles_mgr$add(xf_full_single, xf_full_single)
          self$set_cell_style(sheet, dims, self$styles_mgr$get_xf_id(xf_full_single))
        }

        # create top & bottom piece
        if (nrow(df) >= 2) {

          # top single
          self$styles_mgr$add(top_single, stop_single)
          xf_prev <- get_cell_styles(self, sheet, dim_top_single)
          xf_top_single <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_single))
          self$styles_mgr$add(xf_top_single, xf_top_single)
          self$set_cell_style(sheet, dim_top_single, self$styles_mgr$get_xf_id(xf_top_single))

          # bottom single
          self$styles_mgr$add(bottom_single, sbottom_single)
          xf_prev <- get_cell_styles(self, sheet, dim_bottom_single)
          xf_bottom_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_single))
          self$styles_mgr$add(xf_bottom_single, xf_bottom_single)
          self$set_cell_style(sheet, dim_bottom_single, self$styles_mgr$get_xf_id(xf_bottom_single))
        }

        # create middle piece(s)
        if (nrow(df) >= 3) {

          # middle single
          self$styles_mgr$add(middle_single, smiddle_single)
          xf_prev <- get_cell_styles(self, sheet, dim_middle_single)
          xf_middle_single <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_single))
          self$styles_mgr$add(xf_middle_single, xf_middle_single)
          self$set_cell_style(sheet, dim_middle_single, self$styles_mgr$get_xf_id(xf_middle_single))
        }

      }

      # create left and right single row pieces
      if (ncol(df) >= 2 && nrow(df) == 1) {

        # left single
        self$styles_mgr$add(left_single, sleft_single)
        xf_prev <- get_cell_styles(self, sheet, dim_left_single)
        xf_left_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sleft_single))
        self$styles_mgr$add(xf_left_single, xf_left_single)
        self$set_cell_style(sheet, dim_left_single, self$styles_mgr$get_xf_id(xf_left_single))

        # right single
        self$styles_mgr$add(right_single, sright_single)
        xf_prev <- get_cell_styles(self, sheet, dim_right_single)
        xf_right_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sright_single))
        self$styles_mgr$add(xf_right_single, xf_right_single)
        self$set_cell_style(sheet, dim_right_single, self$styles_mgr$get_xf_id(xf_right_single))

        # add single center piece(s)
        if (ncol(df) >= 3) {

          # center single
          self$styles_mgr$add(center_single, scenter_single)
          xf_prev <- get_cell_styles(self, sheet, dim_center_single)
          xf_center_single <- set_border(xf_prev, self$styles_mgr$get_border_id(scenter_single))
          self$styles_mgr$add(xf_center_single, xf_center_single)
          self$set_cell_style(sheet, dim_center_single, self$styles_mgr$get_xf_id(xf_center_single))
        }

      }

      # create left & right - top & bottom corners pieces
      if (ncol(df) >= 2 && nrow(df) >= 2) {

        # top left
        self$styles_mgr$add(top_left, stop_left)
        xf_prev <- get_cell_styles(self, sheet, dim_top_left)
        xf_top_left <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_left))
        self$styles_mgr$add(xf_top_left, xf_top_left)
        self$set_cell_style(sheet, dim_top_left, self$styles_mgr$get_xf_id(xf_top_left))

        # top right
        self$styles_mgr$add(top_right, stop_right)
        xf_prev <- get_cell_styles(self, sheet, dim_top_right)
        xf_top_right <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_right))
        self$styles_mgr$add(xf_top_right, xf_top_right)
        self$set_cell_style(sheet, dim_top_right, self$styles_mgr$get_xf_id(xf_top_right))

        # bottom left
        self$styles_mgr$add(bottom_left, sbottom_left)
        xf_prev <- get_cell_styles(self, sheet, dim_bottom_left)
        xf_bottom_left <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_left))
        self$styles_mgr$add(xf_bottom_left, xf_bottom_left)
        self$set_cell_style(sheet, dim_bottom_left, self$styles_mgr$get_xf_id(xf_bottom_left))

        # bottom right
        self$styles_mgr$add(bottom_right, sbottom_right)
        xf_prev <- get_cell_styles(self, sheet, dim_bottom_right)
        xf_bottom_right <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_right))
        self$styles_mgr$add(xf_bottom_right, xf_bottom_right)
        self$set_cell_style(sheet, dim_bottom_right, self$styles_mgr$get_xf_id(xf_bottom_right))
      }

      # create left and right middle pieces
      if (ncol(df) >= 2 && nrow(df) >= 3) {

        # middle left
        self$styles_mgr$add(middle_left, smiddle_left)
        xf_prev <- get_cell_styles(self, sheet, dim_middle_left)
        xf_middle_left <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_left))
        self$styles_mgr$add(xf_middle_left, xf_middle_left)
        self$set_cell_style(sheet, dim_middle_left, self$styles_mgr$get_xf_id(xf_middle_left))

        # middle right
        self$styles_mgr$add(middle_right, smiddle_right)
        xf_prev <- get_cell_styles(self, sheet, dim_middle_right)
        xf_middle_right <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_right))
        self$styles_mgr$add(xf_middle_right, xf_middle_right)
        self$set_cell_style(sheet, dim_middle_right, self$styles_mgr$get_xf_id(xf_middle_right))
      }

      # create top and bottom center pieces
      if (ncol(df) >= 3 & nrow(df) >= 2) {

        # top center
        self$styles_mgr$add(top_center, stop_center)
        xf_prev <- get_cell_styles(self, sheet, dim_top_center)
        xf_top_center <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_center))
        self$styles_mgr$add(xf_top_center, xf_top_center)
        self$set_cell_style(sheet, dim_top_center, self$styles_mgr$get_xf_id(xf_top_center))

        # bottom center
        self$styles_mgr$add(bottom_center, sbottom_center)
        xf_prev <- get_cell_styles(self, sheet, dim_bottom_center)
        xf_bottom_center <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_center))
        self$styles_mgr$add(xf_bottom_center, xf_bottom_center)
        self$set_cell_style(sheet, dim_bottom_center, self$styles_mgr$get_xf_id(xf_bottom_center))
      }

      if (nrow(df) > 2 && ncol(df) > 2) {

        # inner cells
        self$styles_mgr$add(inner_cell, sinner_cell)
        xf_prev <- get_cell_styles(self, sheet, dim_inner_cell)
        xf_inner_cell <- set_border(xf_prev, self$styles_mgr$get_border_id(sinner_cell))
        self$styles_mgr$add(xf_inner_cell, xf_inner_cell)
        self$set_cell_style(sheet, dim_inner_cell, self$styles_mgr$get_xf_id(xf_inner_cell))
      }

      invisible(self)
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
      sheet <- private$get_sheet_index(sheet)
      private$do_cell_init(sheet, dims)

      # dim in dataframe can contain various styles. go cell by cell.
      did <- dims_to_dataframe(dims, fill = TRUE)
      # select a few cols and rows to fill
      cols <- (seq_len(ncol(did)) %% every_nth_col) == 0
      rows <- (seq_len(nrow(did)) %% every_nth_row) == 0

      dims <- unname(unlist(did[rows, cols, drop = FALSE]))

      cc <- self$worksheets[[sheet]]$sheet_data$cc
      cc <- cc[cc$r %in% dims, ]
      styles <- unique(cc[["c_s"]])

      for (style in styles) {
        dim <- cc[cc$c_s == style, "r"]

        new_fill <- create_fill(
          gradientFill = gradient_fill,
          patternType = pattern,
          fgColor = color
        )
        self$styles_mgr$add(new_fill, new_fill)

        xf_prev <- get_cell_styles(self, sheet, dim[[1]])
        xf_new_fill <- set_fill(xf_prev, self$styles_mgr$get_fill_id(new_fill))
        self$styles_mgr$add(xf_new_fill, xf_new_fill)
        s_id <- self$styles_mgr$get_xf_id(xf_new_fill)
        self$set_cell_style(sheet, dim, s_id)
      }

      invisible(self)
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
      sheet <- private$get_sheet_index(sheet)
      private$do_cell_init(sheet, dims)

      did <- dims_to_dataframe(dims, fill = TRUE)
      dims <- unname(unlist(did))

      cc <- self$worksheets[[sheet]]$sheet_data$cc
      cc <- cc[cc$r %in% dims, ]
      styles <- unique(cc[["c_s"]])

      for (style in styles) {
        dim <- cc[cc$c_s == style, "r"]

        new_font <- create_font(
          b = bold,
          charset = charset,
          color = color,
          condense = condense,
          extend = extend,
          family = family,
          i = italic,
          name = name,
          outline = outline,
          scheme = scheme,
          shadow = shadow,
          strike = strike,
          sz = size,
          u = underline,
          vertAlign = vertAlign
        )
        self$styles_mgr$add(new_font, new_font)

        xf_prev <- get_cell_styles(self, sheet, dim[[1]])
        xf_new_font <- set_font(xf_prev, self$styles_mgr$get_font_id(new_font))

        self$styles_mgr$add(xf_new_font, xf_new_font)
        s_id <- self$styles_mgr$get_xf_id(xf_new_font)
        self$set_cell_style(sheet, dim, s_id)
      }

      invisible(self)
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
      sheet <- private$get_sheet_index(sheet)
      private$do_cell_init(sheet, dims)

      did <- dims_to_dataframe(dims, fill = TRUE)
      dims <- unname(unlist(did))

      cc <- self$worksheets[[sheet]]$sheet_data$cc
      cc <- cc[cc$r %in% dims, ]
      styles <- unique(cc[["c_s"]])

      if (inherits(numfmt, "character")) {

        for (style in styles) {
          dim <- cc[cc$c_s == style, "r"]

          new_numfmt <- create_numfmt(
            numFmtId = self$styles_mgr$next_numfmt_id(),
            formatCode = numfmt
          )
          self$styles_mgr$add(new_numfmt, new_numfmt)

          xf_prev <- get_cell_styles(self, sheet, dim[[1]])
          xf_new_numfmt <- set_numfmt(xf_prev, self$styles_mgr$get_numfmt_id(new_numfmt))
          self$styles_mgr$add(xf_new_numfmt, xf_new_numfmt)
          s_id <- self$styles_mgr$get_xf_id(xf_new_numfmt)
          self$set_cell_style(sheet, dim, s_id)
        }

      } else { # format is numeric
        for (style in styles) {
          dim <- cc[cc$c_s == style, "r"]
          xf_prev <- get_cell_styles(self, sheet, dim[[1]])
          xf_new_numfmt <- set_numfmt(xf_prev, numfmt)
          self$styles_mgr$add(xf_new_numfmt, xf_new_numfmt)
          s_id <- self$styles_mgr$get_xf_id(xf_new_numfmt)
          self$set_cell_style(sheet, dim, s_id)
        }

      }

      invisible(self)
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
      sheet <- private$get_sheet_index(sheet)
      private$do_cell_init(sheet, dims)

      did <- dims_to_dataframe(dims, fill = TRUE)
      dims <- unname(unlist(did))

      cc <- self$worksheets[[sheet]]$sheet_data$cc
      cc <- cc[cc$r %in% dims, ]
      styles <- unique(cc[["c_s"]])

      for (style in styles) {
        dim <- cc[cc$c_s == style, "r"]
        xf_prev <- get_cell_styles(self, sheet, dim[[1]])
        xf_new_cellstyle <- set_cellstyle(
          xf_node           = xf_prev,
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
        self$styles_mgr$add(xf_new_cellstyle, xf_new_cellstyle)
        s_id <- self$styles_mgr$get_xf_id(xf_new_cellstyle)
        self$set_cell_style(sheet, dim, s_id)
      }

      invisible(self)
    },

    #' @description get sheet style
    #' @param sheet sheet
    #' @param dims dims
    #' @returns a character vector of cell styles
    get_cell_style = function(sheet = current_sheet(), dims) {

      if (length(dims) == 1 && grepl(":", dims))
        dims <- dims_to_dataframe(dims, fill = TRUE)
      sheet <- private$get_sheet_index(sheet)

      # This alters the workbook
      temp <- self$clone()$.__enclos_env__$private$do_cell_init(sheet, dims)

      # if a range is passed (e.g. "A1:B2") we need to get every cell
      dims <- unname(unlist(dims))

      # TODO check that cc$r is alway valid. not sure atm
      sel <- temp$worksheets[[sheet]]$sheet_data$cc$r %in% dims
      temp$worksheets[[sheet]]$sheet_data$cc$c_s[sel]
    },

    #' @description set sheet style
    #' @param sheet sheet
    #' @param dims dims
    #' @param style style
    #' @return The `wbWorksheetObject`, invisibly
    set_cell_style = function(sheet = current_sheet(), dims, style) {

      if (length(dims) == 1 && grepl(":|;", dims))
        dims <- dims_to_dataframe(dims, fill = TRUE)
      sheet <- private$get_sheet_index(sheet)

      private$do_cell_init(sheet, dims)

      # if a range is passed (e.g. "A1:B2") we need to get every cell
      dims <- unname(unlist(dims))

      sel <- self$worksheets[[sheet]]$sheet_data$cc$r %in% dims

      self$worksheets[[sheet]]$sheet_data$cc$c_s[sel] <- style

      invisible(self)
    },

    #' @description clone style from one sheet to another
    #' @param from the worksheet you are cloning
    #' @param to the worksheet the style is applied to
    clone_sheet_style = function(from = current_sheet(), to) {

      id_org <- private$get_sheet_index(from)
      id_new <- private$get_sheet_index(to)

      org_style <- self$worksheets[[id_org]]$sheet_data$cc
      wb_style  <- self$worksheets[[id_new]]$sheet_data$cc

      # only clone styles from sheets with cc
      if (is.null(org_style)) {
        message("'from' has no sheet data styles to clone")
      } else {

        if (is.null(wb_style)) # if null, create empty dataframe
          wb_style <- create_char_dataframe(names(org_style), n = 0)

        # remove all values
        org_style <- org_style[c("r", "row_r", "c_r", "c_s")]

        # do not merge c_s and do not create duplicates
        merged_style <- merge(org_style,
                              wb_style[-which(names(wb_style) == "c_s")],
                              all = TRUE)
        merged_style[is.na(merged_style)] <- ""
        merged_style <- merged_style[!duplicated(merged_style["r"]), ]

        # will be ordere on save
        self$worksheets[[id_new]]$sheet_data$cc <- merged_style

      }


      # copy entire attributes from original sheet to new sheet
      org_rows <- self$worksheets[[id_org]]$sheet_data$row_attr
      new_rows <- self$worksheets[[id_new]]$sheet_data$row_attr

      if (is.null(org_style)) {
        message("'from' has no row styles to clone")
      } else {

        if (is.null(new_rows))
          new_rows <- create_char_dataframe(names(org_rows), n = 0)

        # only add the row information, nothing else
        merged_rows <- merge(org_rows,
                             new_rows["r"],
                             all = TRUE)
        merged_rows[is.na(merged_rows)] <- ""
        merged_rows <- merged_rows[!duplicated(merged_rows["r"]), ]
        ordr <- ordered(order(as.integer(merged_rows$r)))
        merged_rows <- merged_rows[ordr, ]

        self$worksheets[[id_new]]$sheet_data$row_attr <- merged_rows
      }

      self$worksheets[[id_new]]$cols_attr <-
        self$worksheets[[id_org]]$cols_attr

      self$worksheets[[id_new]]$dimension <-
        self$worksheets[[id_org]]$dimension

      self$worksheets[[id_new]]$mergeCells <-
        self$worksheets[[id_org]]$mergeCells

      invisible(self)
    },


    #' @description apply sparkline to worksheet
    #' @param sheet the worksheet you are using
    #' @param sparklines sparkline created by `create_sparkline()`
    add_sparklines = function(
      sheet = current_sheet(),
      sparklines
    ) {
      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$add_sparklines(sparklines)
      invisible(self)
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
      invisible(self)
    },

    get_drawingsref = function() {
      has_drawing <- which(grepl("drawings", self$worksheets_rels))

      rlshp <- NULL
      for (i in has_drawing) {
        rblst <- rbindlist(xml_attr(self$worksheets_rels[[i]], "Relationship"))
        rblst$type <- basename(rblst$Type)
        rblst$id   <- as.integer(gsub("\\D+", "", rblst$Target))
        rblst$sheet <- i

        rlshp <- rbind(rlshp, rblst[rblst$type == "drawing", c("type", "id", "sheet")])
      }

      invisible(rlshp)
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

        if (self$is_chartsheet[i]) {
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
        } ## end of is_chartsheet[i]
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
