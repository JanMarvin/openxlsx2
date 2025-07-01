

# helpers -----------------------------------------------------------------

guard_ws <- function(x) {
  if (grepl(" ", x)) x <- shQuote(x, type = "sh")
  x
}

if_not_missing <- function(x) if (missing(x)) NULL else as.character(x)

# TODO get table Id from table entry
table_ids <- function(wb) {
  z <- 0
  if (!all(identical(unlist(wb$worksheets_rels), character()))) {
    relship <- rbindlist(xml_attr(unlist(wb$worksheets_rels), "Relationship"))
    relship$typ <- basename(relship$Type)
    relship$tid <- as.numeric(gsub("\\D+", "", relship$Target))

    z <- sort(relship$tid[relship$typ == "table"])
  }
  z
}

## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
last_table_id <- function(wb) {
  max(as.integer(rbindlist(xml_attr(wb$tables$tab_xml, "table"))$id), 0)
}

fun_tab_cols <- function(tab_cols) {
  tabCols <- NULL
  for (i in seq_along(tab_cols)) {
    tmp <- xml_node_create(
      "tableColumn",
      xml_attributes = c(id = as.character(i), name = tab_cols[i])
    )
    tabCols <- c(tabCols, tmp)
  }

  xml_node_create(
    "tableColumns",
    xml_attributes = c(count = as.character(length(tabCols))),
    xml_children = tabCols
  )
}

validRow <- function(summary_row) {
  tolower(summary_row) %in% c("above", "below")
}

validCol <- function(summary_col) {
  tolower(summary_col) %in% c("left", "right")
}

lcr <- function(var) {
  # quick function for specifying error message
  paste(var, "must have length 3 where elements correspond to positions: left, center, right.")
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


# R6 class ----------------------------------------------------------------
# Lines 7 and 8 are needed until r-lib/roxygen2#1504 is fixed
#' Workbook class
#'
#' @description
#' This is the class used by `openxlsx2` to modify workbooks from R.
#' You can load an existing workbook with [wb_load()] and create a new one with
#' [wb_workbook()].
#'
#' After that, you can modify the `wbWorkbook` object through two primary methods:
#'
#' *Wrapper Function Method*: Utilizes the `wb` family of functions that support
#'  piping to streamline operations.
#' ``` r
#' wb <- wb_workbook(creator = "My name here") %>%
#'   wb_add_worksheet(sheet = "Expenditure", grid_lines = FALSE) %>%
#'   wb_add_data(x = USPersonalExpenditure, row_names = TRUE)
#' ```
#' *Chaining Method*: Directly modifies the object through a series of chained
#'  function calls.
#' ``` r
#' wb <- wb_workbook(creator = "My name here")$
#'   add_worksheet(sheet = "Expenditure", grid_lines = FALSE)$
#'   add_data(x = USPersonalExpenditure, row_names = TRUE)
#' ```
#'
#' While wrapper functions require explicit assignment of their output to reflect
#' changes, chained functions inherently modify the input object. Both approaches
#' are equally supported, offering flexibility to suit user preferences. The
#' documentation mainly highlights the use of wrapper functions.
#'
#' ``` r
#' # Import workbooks
#' path <- system.file("extdata/openxlsx2_example.xlsx", package = "openxlsx2")
#' wb <- wb_load(path)
#'
#' ## or create one yourself
#' wb <- wb_workbook()
#' # add a worksheet
#' wb$add_worksheet("sheet")
#' # add some data
#' wb$add_data("sheet", cars)
#' # Add data with piping in a different location
#' wb <- wb %>% wb_add_data(x = cars, dims = wb_dims(from_dims = "D4"))
#' # open it in your default spreadsheet software
#' if (interactive()) wb$open()
#' ```
#'
#' Note that the documentation is more complete in each of the wrapper functions.
#' (i.e. `?wb_add_data` rather than `?wbWorkbook`).
#'
#' @param creator character vector of creators. Duplicated are ignored.
#' @param dims Cell range in a sheet
#' @param sheet The name of the sheet
#' @param datetime_created The datetime (as `POSIXt`) the workbook is
#'   created.  Defaults to the current `Sys.time()` when the workbook object
#'   is created, not when the Excel files are saved.
#' @param datetime_modified The datetime (as `POSIXt`) that should be recorded
#'   as last modification date. Defaults to the creation date.
#' @param ... additional arguments
#' @export
wbWorkbook <- R6::R6Class(
  "wbWorkbook",

  # TODO which can be private?

  ## public ----

  public = list(
    #' @field sheet_names The names of the sheets
    sheet_names = character(),

    #' @field calcChain calcChain
    calcChain = character(),

    #' @field charts charts
    charts = list(),

    #' @field is_chartsheet A logical vector identifying if a sheet is a chartsheet.
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

    #' @field core The XML core
    core = character(),

    #' @field custom custom
    custom = character(),

    #' @field drawings drawings
    drawings = NULL,

    #' @field drawings_rels drawings_rels
    drawings_rels = NULL,

    #' @field docMetadata doc_meta_data
    docMetadata = NULL,

    # #' @field drawings_vml drawings_vml
    # drawings_vml = NULL,

    #' @field activeX activeX
    activeX = NULL,

    #' @field embeddings embeddings
    embeddings = NULL,

    #' @field externalLinks externalLinks
    externalLinks = NULL,

    #' @field externalLinksRels externalLinksRels
    externalLinksRels = NULL,

    #' @field featurePropertyBag featurePropertyBag
    featurePropertyBag = NULL,

    #' @field headFoot The header and footer
    headFoot = NULL,

    #' @field media media
    media = NULL,

    #' @field metadata contains cell/value metadata imported on load from xl/metadata.xml
    metadata = NULL,

    #' @field persons Persons of the workbook. to be used with [wb_add_thread()]
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

    #' @field richData richData
    richData = NULL,

    #' @field slicers slicers
    slicers = NULL,

    #' @field slicerCaches slicerCaches
    slicerCaches = NULL,

    #' @field sharedStrings sharedStrings
    sharedStrings = structure(list(), uniqueCount = 0L),

    #' @field styles_mgr styles_mgr
    styles_mgr = NULL,

    #' @field tables tables
    tables = NULL,

    #' @field tables.xml.rels tables.xml.rels
    tables.xml.rels = NULL,

    #' @field theme theme
    theme = NULL,

    #' @field vbaProject vbaProject
    vbaProject = NULL,

    #' @field vml vml
    vml = NULL,

    #' @field vml_rels vml_rels
    vml_rels = NULL,

    #' @field comments Comments (notes) present in the workbook.
    comments = list(),

    #' @field threadComments Threaded comments
    threadComments = NULL,

    #' @field timelines timelines
    timelines = NULL,

    #' @field timelineCaches timelineCaches
    timelineCaches = NULL,

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

    #' @field namedSheetViews namedSheetViews
    namedSheetViews = character(),

    #' @description
    #' Creates a new `wbWorkbook` object
    #' @param title,subject,category,keywords,comments,manager,company workbook properties
    #' @param theme Optional theme identified by string or number
    #' @param ... additional arguments
    #' @return a `wbWorkbook` object
    initialize = function(
      creator           = NULL,
      title             = NULL,
      subject           = NULL,
      category          = NULL,
      datetime_created  = Sys.time(),
      datetime_modified = NULL,
      theme             = NULL,
      keywords          = NULL,
      comments          = NULL,
      manager           = NULL,
      company           = NULL,
      ...
    ) {

      force(datetime_created)

      standardize_case_names(...)

      self$app <- genBaseApp()
      self$charts <- list()
      self$is_chartsheet <- logical()

      self$connections <- NULL
      self$Content_Types <- genBaseContent_Type()

      creator <- creator %||%
        getOption("openxlsx2.creator",
                  default = Sys.getenv("USERNAME", unset = Sys.getenv("USER")))
        # USERNAME is present for (Windows, Linux) "USER" is present for Mac

      # Internal option to alleviate timing problems in CI and CRAN
      datetime_created <- getOption("openxlsx2.datetimeCreated", datetime_created)


      assert_class(creator,          "character")
      assert_class(title,            "character", or_null = TRUE)
      assert_class(subject,          "character", or_null = TRUE)
      assert_class(category,         "character", or_null = TRUE)
      assert_class(keywords,         "character", or_null = TRUE)
      assert_class(comments,         "character", or_null = TRUE)
      assert_class(manager,          "character", or_null = TRUE)
      assert_class(company,          "character", or_null = TRUE)

      assert_class(datetime_created, "POSIXt")
      assert_class(datetime_modified, "POSIXt", or_null = TRUE)

      # Avoid modtime being slightly different from createtime by two distinct
      # Sys.time() calls
      if (is.null(datetime_modified)) datetime_modified <- datetime_created

      stopifnot(
        length(title) <= 1L,
        length(category) <= 1L,
        length(datetime_created) == 1L
      )

      self$set_properties(
        creator            = creator,
        title              = title,
        subject            = subject,
        category           = category,
        datetime_created   = datetime_created,
        datetime_modified  = datetime_modified,
        keywords           = keywords,
        comments           = comments,
        manager            = manager,
        company            = company
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

      self$richData <- NULL

      self$sheet_names <- character()
      self$sheetOrder <- integer()

      self$sharedStrings <- list()
      attr(self$sharedStrings, "uniqueCount") <- 0

      # add styles_mgr and set default styles. will initialize after theme
      self$styles_mgr <- style_mgr$new(self)
      self$styles_mgr$styles <- genBaseStyleSheet()

      empty_cellXfs <- data.frame(
        numFmtId = "0",
        fontId   = "0",
        fillId   = "0",
        borderId = "0",
        xfId     = "0",
        stringsAsFactors = FALSE
      )
      self$styles_mgr$styles$cellXfs <- write_xf(empty_cellXfs)

      self$tables <- NULL
      self$tables.xml.rels <- NULL

      if (is.null(theme)) {
        self$theme <- NULL
      } else {
        # read predefined themes
        thm_rds <- system.file("extdata", "themes.rds", package = "openxlsx2")
        themes <- readRDS(thm_rds)

        if (is.character(theme)) {
          sel <- match(theme, names(themes))
          err <- is.na(sel)
        } else {
          sel <- theme
          err <- sel > length(themes)
        }

        if (err) {
          message("theme ", theme, " not found falling back to default theme")
        } else {
          self$theme <- stringi::stri_unescape_unicode(themes[[sel]])

          # create the default font for the style
          font_scheme <- xml_node(self$theme, "a:theme", "a:themeElements", "a:fontScheme")
          minor_font <- xml_attr(font_scheme, "a:fontScheme", "a:minorFont", "a:latin")[[1]][["typeface"]]

          self$styles_mgr$styles$fonts <- create_font(
            sz = 11,
            color = wb_color(theme = 1),
            name = minor_font,
            family = "2",
            scheme = "minor"
          )

        }
      }

      self$styles_mgr$initialize(self)

      self$vbaProject <- NULL
      self$vml <- NULL
      self$vml_rels <- NULL

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
    #' @return The integer position of the sheet
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
    #' @param tab_color tab_color
    #' @param zoom zoom
    #' @param visible visible
    #' @return The `wbWorkbook` object, invisibly
    add_chartsheet = function(
      sheet     = next_sheet(),
      tab_color = NULL,
      zoom      = 100,
      visible   = c("true", "false", "hidden", "visible", "veryhidden"),
      ...
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
        default_sheet_name <- getOption("openxlsx2.sheet.default_name", "Sheet ")
        default_sheets <- self$sheet_names[grepl(default_sheet_name, self$sheet_names)]
        max_sheet_num <- max(
          0,
          as.integer(gsub("\\D+", "", default_sheets))
        )
        sheet <- paste0(
          default_sheet_name,
          max_sheet_num + 1L
        )
      }

      sheet <- as.character(sheet)
      private$validate_new_sheet(sheet)
      sheet_name <- replace_legal_chars(sheet)


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

      standardize(...)

      if (!is.null(tab_color) && !is_wbColour(tab_color)) {
        validate_color(tab_color, msg = "Invalid tab_color in add_chartsheet.")
        tabColor <- wb_color(tab_color)
      } else {
        tabColor <- tab_color
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
          tab_color = tabColor
        )
      )

      self$worksheets[[newSheetIndex]]$set_sheetview(
        workbook_view_id = 0,
        zoom_scale       = zoom,
        tab_selected     = newSheetIndex == 1
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

      ## create sheet.rels to simplify id assignment
      # new_drawings_idx <- length(self$drawings) + 1
      # self$drawings[[new_drawings_idx]]      <- ""
      # self$drawings_rels[[new_drawings_idx]] <- ""

      self$worksheets_rels[[newSheetIndex]]  <- genBaseSheetRels(newSheetIndex)
      self$is_chartsheet[[newSheetIndex]]    <- TRUE
      # self$vml_rels[[newSheetIndex]]         <- list()
      # self$vml[[newSheetIndex]]              <- list()

      self$append("sheetOrder", newSheetIndex)
      private$set_single_sheet_name(newSheetIndex, sheet_name, sheet)

      invisible(self)
    },

    #' @description
    #' Add worksheet to the `wbWorkbook` object
    #' @param grid_lines gridLines
    #' @param row_col_headers rowColHeaders
    #' @param tab_color tabColor
    #' @param zoom zoom
    #' @param header header
    #' @param footer footer
    #' @param odd_header oddHeader
    #' @param odd_footer oddFooter
    #' @param even_header evenHeader
    #' @param even_footer evenFooter
    #' @param first_header firstHeader
    #' @param first_footer firstFooter
    #' @param visible visible
    #' @param has_drawing hasDrawing
    #' @param paper_size paperSize
    #' @param orientation orientation
    #' @param hdpi hdpi
    #' @param vdpi vdpi
    #' @return The `wbWorkbook` object, invisibly
    add_worksheet = function(
      sheet           = next_sheet(),
      grid_lines      = TRUE,
      row_col_headers = TRUE,
      tab_color       = NULL,
      zoom            = 100,
      header          = NULL,
      footer          = NULL,
      odd_header      = header,
      odd_footer      = footer,
      even_header     = header,
      even_footer     = footer,
      first_header    = header,
      first_footer    = footer,
      visible         = c("true", "false", "hidden", "visible", "veryhidden"),
      has_drawing     = FALSE,
      paper_size      = getOption("openxlsx2.paperSize", default = 9),
      orientation     = getOption("openxlsx2.orientation", default = "portrait"),
      hdpi            = getOption("openxlsx2.hdpi", default = getOption("openxlsx2.dpi", default = 300)),
      vdpi            = getOption("openxlsx2.vdpi", default = getOption("openxlsx2.dpi", default = 300)),
      ...
    ) {

      standardize(...)

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
        default_sheet_name <- getOption("openxlsx2.sheet.default_name", "Sheet ")
        default_sheets <- self$sheet_names[grepl(default_sheet_name, self$sheet_names)]
        max_sheet_num <- max(
          0,
          as.integer(gsub("\\D+", "", default_sheets))
        )
        sheet <- paste0(
          default_sheet_name,
          max_sheet_num + 1L
        )
      }

      sheet <- as.character(sheet)
      private$validate_new_sheet(sheet)
      sheet_name <- replace_legal_chars(sheet)

      if (!is.logical(grid_lines) | length(grid_lines) > 1) {
        fail <- TRUE
        msg <- c(msg, "grid_lines must be a logical of length 1.")
      }

      if (!is.null(tab_color) && !is_wbColour(tab_color)) {
        validate_color(tab_color, msg = "Invalid tab_color in add_worksheet.")
        tabColor <- wb_color(tab_color)
      } else {
        tabColor <- tab_color
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

      if (!is.null(odd_header) & length(odd_header) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("header"))
      }

      if (!is.null(odd_footer) & length(odd_footer) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("footer"))
      }

      if (!is.null(even_header) & length(even_header) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("evenHeader"))
      }

      if (!is.null(even_footer) & length(even_footer) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("evenFooter"))
      }

      if (!is.null(first_header) & length(first_header) != 3) {
        fail <- TRUE
        msg <- c(msg, lcr("firstHeader"))
      }

      if (!is.null(first_footer) & length(first_footer) != 3) {
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
          tab_color    = tabColor,
          odd_header   = odd_header,
          odd_footer   = odd_footer,
          even_header  = even_header,
          even_footer  = even_footer,
          first_header = first_header,
          first_footer = first_footer,
          paper_size   = paper_size,
          orientation  = orientation,
          hdpi         = hdpi,
          vdpi         = vdpi,
          grid_lines   = grid_lines
        )
      )

      # NULL or TRUE/FALSE
      rightToLeft <- getOption("openxlsx2.rightToLeft")

      # set preselected set for sheetview
      self$worksheets[[newSheetIndex]]$set_sheetview(
        workbook_view_id     = 0,
        zoom_scale           = zoom,
        show_grid_lines      = grid_lines,
        show_row_col_headers = row_col_headers,
        tab_selected         = newSheetIndex == 1,
        right_to_left        = rightToLeft
      )


      ## update content_tyes
      ## add a drawing.xml for the worksheet
      if (has_drawing) {
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
      # self$drawings[[new_drawings_idx]]      <- ""
      # self$drawings_rels[[new_drawings_idx]] <- ""

      self$worksheets_rels[[newSheetIndex]]  <- genBaseSheetRels(newSheetIndex)
      # self$vml_rels[[newSheetIndex]]         <- list()
      # self$vml[[newSheetIndex]]              <- list()
      self$is_chartsheet[[newSheetIndex]]    <- FALSE
      # self$comments[[newSheetIndex]]         <- list()
      # self$threadComments[[newSheetIndex]]   <- list()

      self$append("sheetOrder", as.integer(newSheetIndex))
      private$set_single_sheet_name(newSheetIndex, sheet_name, sheet)

      invisible(self)
    },


    #' @description
    #' Clone a workbooksheet to another workbook
    #' @param old name of worksheet to clone
    #' @param new name of new worksheet to add
    #' @param from name of new worksheet to add
    clone_worksheet = function(old = current_sheet(), new = next_sheet(), from = NULL) {

      if (is.null(from)) {
        from        <- self$clone()
        external_wb <- FALSE
        suffix      <- "_n"
      } else {
        external_wb <- TRUE
        suffix      <- ""
        assert_workbook(from)
      }

      sheet <- new
      private$validate_new_sheet(sheet)
      new <- sheet

      old <- from$.__enclos_env__$private$get_sheet_index(old)

      newSheetIndex <- length(self$worksheets) + 1L
      private$set_current_sheet(newSheetIndex)
      sheetId <- private$get_sheet_id_max() # checks for length of worksheets

      if (!all(from$charts$chartEx == "")) {
        warning(
          "The file you have loaded contains chart extensions. At the moment,",
          " cloning worksheets can damage the output."
        )
      }

      # not the best but a quick fix
      new_raw <- new
      new <- replace_legal_chars(new)

      ## copy visibility from cloned sheet!
      visible <- rbindlist(xml_attr(from$workbook$sheets[[old]], "sheet"))$state

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
      self$append("worksheets", from$worksheets[[old]]$clone(deep = TRUE))

      ## update content_tyes
      ## add a drawing.xml for the worksheet
      # FIXME only add what is needed. If no previous drawing is found, don't
      # add a new one
      self$append("Content_Types", c(
        if (from$is_chartsheet[old]) {
          sprintf('<Override PartName="/xl/chartsheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>', newSheetIndex)
        } else {
          sprintf('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', newSheetIndex)
        }
      ))

      ## Update xl/rels
      self$append(
        "workbook.xml.rels",
        if (from$is_chartsheet[old]) {
          sprintf('<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet%s.xml"/>', newSheetIndex)
        } else {
          sprintf('<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>', newSheetIndex)
        }
      )

      ## create sheet.rels to simplify id assignment
      self$worksheets_rels[[newSheetIndex]] <- from$worksheets_rels[[old]]

      ## TODO actually check this and add a similar warning for embeddings
      ## TODO it should be able to clone this
      if (any(grepl("activeX", self$worksheets_rels[[newSheetIndex]]))) {
        warning("The cloned sheet contains an activeX element. Cloning this is not yet handled.")
      }

      new_drawing_sheet <- NULL
      if (length(from$worksheets[[old]]$relships$drawing)) {

        drawing_id <- from$worksheets[[old]]$relships$drawing

        new_drawing_sheet <- length(self$drawings) + 1L
        new_drawing_rels  <- length(self$drawings_rels) + 1L

        # if drawings_rels is list(), appending will create multiple lists
        self$append("drawings_rels", list(from$drawings_rels[[drawing_id]]))

        # select the latest addition to drawings_rels
        drawings_rels <- self$drawings_rels[[new_drawing_rels]]

        # For charts we have to modify the name of the chart in the xml code
        # give each chart its own filename (images can re-use the same file, but charts can't)
        for (dl in seq_along(drawings_rels)) {
          chartfiles <- reg_match(drawings_rels[dl], "(?<=charts/)chart[0-9]+\\.xml")

          for (cf in chartfiles) {
            chartid <- NROW(self$charts) + 1L
            newname <- stringi::stri_join("chart", chartid, ".xml")
            old_chart <- as.integer(gsub("\\D+", "", cf))
            self$charts <- rbind(self$charts, from$charts[old_chart, ])

            # Read the chartfile and adjust all formulas to point to the new
            # sheet name instead of the clone source

            chart <- self$charts$chart[chartid]
            self$charts$rels[chartid] <- gsub(
              "?drawing[0-9]+.xml",
              paste0("drawing", chartid, ".xml"),
              self$charts$rels[chartid]
            )

            old_sheet_name <- guard_ws(from$sheet_names[[old]])
            new_sheet_name <- guard_ws(new)

            ## we need to replace "'oldname'" as well as "oldname"
            chart <- gsub(
              paste0(">", old_sheet_name, "!"),
              paste0(">", new_sheet_name, "!"),
              chart,
              perl = TRUE
            )

            self$charts$chart[chartid] <- chart

            # two charts can not point to the same rels
            if (self$charts$rels[chartid] != "") {
              self$charts$rels[chartid] <- gsub(
                stringi::stri_join(old_chart, ".xml"),
                stringi::stri_join(chartid, ".xml"),
                self$charts$rels[chartid]
              )
            }

            drawings_rels[dl] <- gsub(stringi::stri_join("(?<=charts/)", cf), newname, drawings_rels[dl], perl = TRUE)
          }
        }

        self$drawings_rels[[new_drawing_rels]] <- drawings_rels

        self$append("drawings", from$drawings[[drawing_id]])
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

        old_s_id <- from$worksheets[[old]]$relships$slicer

        cloned_slicers <- from$slicers[[old_s_id]]
        slicer_attr <- xml_attr(cloned_slicers, "slicers")

        # Replace name with name_n. This will prevent the slicer from loading,
        # but the xlsx file is not broken
        slicer_child <- xml_node(cloned_slicers, "slicers", "slicer")
        slicer_df <- rbindlist(xml_attr(slicer_child, "slicer"))[c("name", "cache", "caption", "rowHeight")]
        slicer_df$name <- paste0(slicer_df$name, suffix)
        slicer_child <- df_to_xml("slicer", slicer_df)

        self$slicers[[newid]] <- xml_node_create("slicers", slicer_child, slicer_attr[[1]])

        self$worksheets[[newSheetIndex]]$relships$slicer <- newid

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

      rid <- as.integer(sub("\\D+", "", get_relship_id(obj = self$worksheets_rels[[newSheetIndex]], "timeline")))
      if (length(rid)) {

        warning("Cloning timelines is not yet supported. It will not appear on the sheet.")
        self$worksheets_rels[[newSheetIndex]] <- relship_no(obj = self$worksheets_rels[[newSheetIndex]], x = "timeline")

        newid <- length(self$timelines) + 1L

        old_t_id <- from$worksheets[[old]]$relships$timeline

        cloned_timelines <- from$timelines[[old_t_id]]
        timeline_attr <- xml_attr(cloned_timelines, "timelines")

        # Replace name with name_n. This will prevent the timeline from loading,
        # but the xlsx file is not broken
        timeline_child <- xml_node(cloned_timelines, "timelines", "timeline")
        timeline_df <- rbindlist(xml_attr(timeline_child, "timeline"))[c("name", "xr10:uid", "cache", "caption", "level", "selectionLevel", "scrollPosition")]
        timeline_df$name <- paste0(timeline_df$name, suffix)
        timeline_child <- df_to_xml("timeline", timeline_df)

        self$timelines[[newid]] <- xml_node_create("timelines", timeline_child, timeline_attr[[1]])

        self$worksheets[[newSheetIndex]]$relships$timeline <- newid

        self$worksheets_rels[[newSheetIndex]] <- c(
          self$worksheets_rels[[newSheetIndex]],
          sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.microsoft.com/office/2011/relationships/timeline\" Target=\"../timelines/timeline%s.xml\"/>",
                  rid,
                  newid)
        )

        self$Content_Types <- c(
          self$Content_Types,
          sprintf("<Override PartName=\"/xl/timelines/timeline%s.xml\" ContentType=\"application/vnd.ms-excel.timeline+xml\"/>", newid)
        )
      }

      if (!is.null(self$richData)) {
        warning("Cloning richData (e.g., cells with picture) is not yet supported. The output file will be broken.")
      }

      # The IDs in the drawings array are sheet-specific, so within the new
      # cloned sheet the same IDs can be used => no need to modify drawings
      vml_id <- from$worksheets[[old]]$relships$vmlDrawing
      cmt_id <- from$worksheets[[old]]$relships$comments
      trd_id <- from$worksheets[[old]]$relships$threadedComment

      if (length(vml_id)) {
        self$append("vml",      from$vml[[vml_id]])
        self$append("vml_rels", from$vml_rels[[vml_id]])
        self$worksheets[[newSheetIndex]]$relships$vmlDrawing <- length(self$vml)
      }

      if (length(cmt_id)) {
        self$append("comments", from$comments[cmt_id])
        self$worksheets[[newSheetIndex]]$relships$comments <- length(self$comments)
      }

      if (length(trd_id)) {
        self$append("threadComments", from$threadComments[cmt_id])
        self$worksheets[[newSheetIndex]]$relships$threadedComment <- length(self$threadComments)
      }

      self$is_chartsheet[[newSheetIndex]]  <- from$is_chartsheet[[old]]

      self$append("sheetOrder", as.integer(newSheetIndex))
      self$append("sheet_names", new)
      private$set_single_sheet_name(pos = newSheetIndex, clean = new, raw = new_raw)


      ############################
      ## DRAWINGS

      # if we have drawings to clone, remove every table reference from Relationship

      rid <- as.integer(sub("\\D+", "", get_relship_id(obj = self$worksheets_rels[[newSheetIndex]], x = "drawing")))

      if (length(rid) && !is.null(new_drawing_sheet)) {

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

      ## TODO need to collect table dxfs styles, apply them to the workbook
      ## and update the table.xml file with the new dxfs ids. Maybe we can
      ## set these to the default value 0 to avoid broken spreadsheets

      # if we have tables to clone, remove every table referece from Relationship
      rid <- as.integer(sub("\\D+", "", get_relship_id(obj = self$worksheets_rels[[newSheetIndex]], x = "table")))

      if (length(rid)) {

        self$worksheets_rels[[newSheetIndex]] <- relship_no(obj = self$worksheets_rels[[newSheetIndex]], x = "table")

        # make this the new sheets object
        tbls <- from$tables[from$tables$tab_sheet == old, ]
        if (NROW(tbls)) {

          # newid and rid can be different. ids must be unique
          if (!is.null(self$tables$tab_xml))
            newid <- max(as.integer(rbindlist(xml_attr(self$tables$tab_xml, "table"))$id)) + seq_along(rid)
          else
            newid <- 1L

          if (any(stringi::stri_join(tbls$tab_name, suffix) %in% self$tables$tab_name)) {
            tbls$tab_name <- stringi::stri_join(tbls$tab_name, "1")
          }

          # add _n to all table names found
          tbls$tab_name <- stringi::stri_join(tbls$tab_name, suffix)
          tbls$tab_sheet <- newSheetIndex
          # modify tab_xml with updated name, displayName and id
          tbls$tab_xml <- vapply(
            seq_len(nrow(tbls)),
            function(x) {
              xml_attr_mod(
                tbls$tab_xml[x],
                xml_attributes = c(
                  name = tbls$tab_name[x],
                  displayName = tbls$tab_name[x],
                  id = newid[x]
                )
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

      if (external_wb) {

        if (length(from$media)) {

          # get old drawing id, must not match new drawing id
          old_drawing_sheet <- from$worksheets[[old]]$relships$drawing

          if (length(old_drawing_sheet)) {

            # assuming that if drawing was copied, this is the new drawing id
            new_drawing_sheet <- length(self$drawings)
            self$worksheets[[newSheetIndex]]$relships$drawing <- new_drawing_sheet

            # we pick up the drawing relationship. This is something like: "../media/image1.jpg"
            # because we might end up with multiple files with similar names, we have to rename
            # the media file and update the drawing relationship
            # TODO has every drawing a drawing_rel of the same size?
            if (all(nchar(self$drawings_rels[[new_drawing_rels]]))) {

              drels <- rbindlist(xml_attr(self$drawings_rels[[new_drawing_rels]], "Relationship"))
              fe <- unique(file_ext2(drels$Target))

              cte <- sprintf("<Default Extension=\"%s\" ContentType=\"image/%s\"/>", fe, fe)
              sel <- which(!cte %in% self$Content_Types)

              if (length(sel)) {
                self$append("Content_Types", sprintf("<Default Extension=\"%s\" ContentType=\"image/%s\"/>", fe, fe))
              }

              if (ncol(drels) && any(basename(drels$Type) == "image")) {
                sel <- basename(drels$Type) == "image"
                targets <- basename2(drels$Target)[sel]
                media_names <- from$media[targets %in% names(from$media)]

                onams    <- names(media_names)
                mnams    <- vector("character", length(onams))
                next_ids <- length(names(self$media)) + seq_along(mnams)

                # we might have multiple media references on a sheet
                for (i in seq_along(onams)) {
                  media_id   <- as.integer(gsub("\\D+", "", onams[i]))
                  # take filetype + number + file extension
                  # e.g. "image5.jpg" and return "image2.jpg"
                  mnams[i] <- gsub("(\\d+)\\.(\\w+)", paste0(next_ids[i], ".\\2"), onams[i])
                }
                names(media_names) <- mnams

                # update relationship
                self$drawings_rels[[new_drawing_rels]] <- stringi::stri_replace_all_fixed(
                  self$drawings_rels[[new_drawing_rels]],
                  pattern = onams,
                  replacement = mnams,
                  vectorize_all = FALSE
                )

                # append media
                self$append("media", media_names)
              }
            }
          }
        }


        wrels <- rbindlist(xml_attr(self$worksheets_rels[[newSheetIndex]], "Relationship"))
        if (ncol(wrels) && any(sel <- basename(wrels$Type) == "pivotTable")) {
          ## Need to collect the pivot table xml below, apply it to the workbook
          ## and update the references with the new IDs
          # pt <- which(sel)
          # self$pivotTables          <- from$pivotTables[pt]
          # self$pivotTables.xml.rels <- from$pivotTables.xml.rels[pt]
          # self$pivotDefinitions     <- from$pivotDefinitions[pt]
          # self$pivotDefinitionsRels <- from$pivotDefinitionsRels[pt]
          # self$pivotRecords         <- from$pivotRecords[pt]
          #
          # self$append(
          #   "workbook.xml.rels",
          #   "<Relationship Id=\"rId20001\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition1.xml\"/>"
          # )
          #
          # self$append(
          #   "Content_Types",
          #   c(
          #     "<Override PartName=\"/xl/pivotTables/pivotTable1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"/>",
          #     "<Override PartName=\"/xl/pivotCache/pivotCacheDefinition1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/>",
          #     "<Override PartName=\"/xl/pivotCache/pivotCacheRecords1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml\"/>"
          #   )
          # )
          #
          # self$workbook$pivotCaches <- "<pivotCaches><pivotCache cacheId=\"0\" r:id=\"rId20001\"/></pivotCaches>"
          # self$styles_mgr$styles$dxfs         <- from$styles_mgr$styles$dxfs
          # self$styles_mgr$styles$cellStyles   <- from$styles_mgr$styles$cellStyles
          # self$styles_mgr$styles$cellStyleXfs <- from$styles_mgr$styles$cellStyleXfs

          warning("Cloning pivot tables over workbooks is not yet supported.")
        }

        # update sheet styles
        style   <- get_cellstyle(from, sheet = old)
        # only if styles are present
        if (!is.null(style)) {
          new_sty <- set_cellstyles(self, style = style)
          new_s   <- unname(new_sty[match(self$worksheets[[newSheetIndex]]$sheet_data$cc$c_s, names(new_sty))])
          new_s[is.na(new_s)] <- ""
          self$worksheets[[newSheetIndex]]$sheet_data$cc$c_s <- new_s
          rm(style, new_s, new_sty)
        }

        style   <- get_colstyle(from, sheet = old)
        # only if styles are present
        if (!is.null(style)) {
          new_sty <- set_cellstyles(self, style = style)
          cols    <- self$worksheets[[newSheetIndex]]$unfold_cols()
          new_s   <- unname(new_sty[match(cols$style, names(new_sty))])
          new_s[is.na(new_s)] <- ""
          cols$style <- new_s
          self$worksheets[[newSheetIndex]]$fold_cols(cols)
          rm(style, new_s, new_sty)
        }

        style   <- get_rowstyle(from, sheet = old)
        # only if styles are present
        if (!is.null(style)) {
          new_sty <- set_cellstyles(self, style = style)
          new_s   <- unname(new_sty[match(self$worksheets[[newSheetIndex]]$sheet_data$row_attr$s, names(new_sty))])
          new_s[is.na(new_s)] <- ""
          self$worksheets[[newSheetIndex]]$sheet_data$row_attr$s <- new_s
          rm(style, new_s, new_sty)
        }

        # TODO dxfs styles for (pivot) table styles and conditional formatting
        if (length(from$styles_mgr$get_dxf())) {
          msg <- "Input file has dxf styles. These are not cloned. Some styles might be broken and spreadsheet software might complain."
          warning(msg, call. = FALSE)
        }

        clone_shared_strings(from, old, self, newSheetIndex)
      }

      invisible(self)
    },

    ### add data ----

    #' @description add data
    #' @param x x
    #' @param start_col startCol
    #' @param start_row startRow
    #' @param array array
    #' @param col_names colNames
    #' @param row_names rowNames
    #' @param with_filter withFilter
    #' @param name name
    #' @param sep sep
    #' @param apply_cell_style applyCellStyle
    #' @param remove_cell_style if writing into existing cells, should the cell style be removed?
    #' @param na.strings Value used for replacing `NA` values from `x`. Default
    #'   `na_strings()` uses the special `#N/A` value within the workbook.
    #' @param inline_strings write characters as inline strings
    #' @param enforce enforce that selected dims is filled. For this to work, `dims` must match `x`
    #' @param return The `wbWorkbook` object
    add_data = function(
        sheet            = current_sheet(),
        x,
        dims              = wb_dims(start_row, start_col),
        start_col         = 1,
        start_row         = 1,
        array             = FALSE,
        col_names         = TRUE,
        row_names         = FALSE,
        with_filter       = FALSE,
        name              = NULL,
        sep               = ", ",
        apply_cell_style  = TRUE,
        remove_cell_style = FALSE,
        na.strings        = na_strings(),
        inline_strings    = TRUE,
        enforce           = FALSE,
        ...
      ) {

      standardize(...)
      if (missing(x)) stop("`x` is missing")
      if (length(self$sheet_names) == 0) {
        stop(
          "Can't add data to a workbook with no worksheet.\n",
          "Did you forget to add a worksheet with `wb_add_worksheet()`?",
          call. = FALSE
          )
      }

      do_write_data(
        wb                = self,
        sheet             = sheet,
        x                 = x,
        dims              = dims,
        start_col         = start_col,
        start_row         = start_row,
        array             = array,
        col_names         = col_names,
        row_names         = row_names,
        with_filter       = with_filter,
        name              = name,
        sep               = sep,
        apply_cell_style  = apply_cell_style,
        remove_cell_style = remove_cell_style,
        na.strings        = na.strings,
        inline_strings    = inline_strings,
        enforce           = enforce
      )
      invisible(self)
    },

    #' @description add a data table
    #' @param x x
    #' @param start_col startCol
    #' @param start_row startRow
    #' @param col_names colNames
    #' @param row_names rowNames
    #' @param table_style tableStyle
    #' @param table_name tableName
    #' @param with_filter withFilter
    #' @param sep sep
    #' @param first_column firstColumn
    #' @param last_column lastColumn
    #' @param banded_rows bandedRows
    #' @param banded_cols bandedCols
    #' @param apply_cell_style applyCellStyle
    #' @param remove_cell_style if writing into existing cells, should the cell style be removed?
    #' @param na.strings Value used for replacing `NA` values from `x`. Default
    #'   `na_strings()` uses the special `#N/A` value within the workbook.
    #' @param inline_strings write characters as inline strings
    #' @param total_row write total rows to table
    #' @param ... additional arguments
    #' @return The `wbWorkbook` object
    add_data_table = function(
        sheet             = current_sheet(),
        x,
        dims              = wb_dims(start_row, start_col),
        start_col         = 1,
        start_row         = 1,
        col_names         = TRUE,
        row_names         = FALSE,
        table_style       = "TableStyleLight9",
        table_name        = NULL,
        with_filter       = TRUE,
        sep               = ", ",
        first_column      = FALSE,
        last_column       = FALSE,
        banded_rows       = TRUE,
        banded_cols       = FALSE,
        apply_cell_style  = TRUE,
        remove_cell_style = FALSE,
        na.strings        = na_strings(),
        inline_strings    = TRUE,
        total_row         = FALSE,
        ...
    ) {

      standardize(...)
      if (missing(x)) stop("`x` is missing")
      if (length(self$sheet_names) == 0) {
        stop(
          "Can't add data to a workbook with no worksheet.\n",
          "Did you forget to add a worksheet with `wb_add_worksheet()`?",
          call. = FALSE
        )
      }

      do_write_datatable(
        wb                = self,
        x                 = x,
        sheet             = sheet,
        dims              = dims,
        start_col         = start_col,
        start_row         = start_row,
        col_names         = col_names,
        row_names         = row_names,
        table_style       = table_style,
        table_name        = table_name,
        with_filter       = with_filter,
        sep               = sep,
        first_column      = first_column,
        last_column       = last_column,
        banded_rows       = banded_rows,
        banded_cols       = banded_cols,
        apply_cell_style  = apply_cell_style,
        remove_cell_style = remove_cell_style,
        na.strings        = na.strings,
        inline_strings    = inline_strings,
        total_row         = total_row
      )
      invisible(self)
    },


    #' @description add pivot table
    #' @param x a wb_data object
    #' @param dims the worksheet cell where the pivot table is placed
    #' @param filter a character object with names used to filter
    #' @param rows a character object with names used as rows
    #' @param cols a character object with names used as cols
    #' @param data a character object with names used as data
    #' @param fun a character object of functions to be used with the data
    #' @param params a list of parameters to modify pivot table creation
    #' @param pivot_table a character object with a name for the pivot table
    #' @param slicer a character object with names used as slicer
    #' @param timeline a character object with names used as timeline
    #' @details
    #' `fun` can be either of AVERAGE, COUNT, COUNTA, MAX, MIN, PRODUCT, STDEV,
    #' STDEVP, SUM, VAR, VARP
    #' @return The `wbWorkbook` object
    add_pivot_table = function(
      x,
      sheet = next_sheet(),
      dims = "A3",
      filter,
      rows,
      cols,
      data,
      fun,
      params,
      pivot_table,
      slicer,
      timeline
    ) {

      if (missing(x))
        stop("x cannot be missing in add_pivot_table")

      assert_class(x, "wb_data")
      add_sheet <- is_waiver(sheet) && sheet == "next_sheet"
      sheet <- private$get_sheet_index(sheet)

      if (missing(filter))      filter <- substitute()
      if (missing(rows))        rows   <- substitute()
      if (missing(cols))        cols   <- substitute()
      if (missing(data))        data   <- substitute()
      if (missing(fun))         fun    <- substitute()
      if (missing(pivot_table)) pivot_table <- NULL
      if (missing(params))      params <- NULL

      if (anyDuplicated(c(if_not_missing(filter), if_not_missing(rows), if_not_missing(cols)))) {
        stop("duplicated variable in filter, rows, and cols detected.")
      }

      if (!missing(fun) && !missing(data)) {
        if (length(fun) < length(data)) {
          fun <- rep(fun[1], length(data))
        }
      }

      if (any(sel <- duplicated(tolower(names(x))))) {
        nms <- names(x)
        names(x) <- fix_pt_names(nms)
      }

      # for now we use a new worksheet
      if (add_sheet) {
        self$add_worksheet()
      }

      numfmts <- NULL
      if (!is.null(numfmt <- params$numfmt)) {
        if (length(numfmt) != length(data))
          stop("length of numfmt and data does not match")

        for (i in seq_along(numfmt)) {
          if (names(numfmt)[i] == "formatCode") {
            numfmt_i <- self$styles_mgr$next_numfmt_id()
            sty_i <- create_numfmt(numfmt_i, formatCode = numfmt[i])
            self$add_style(sty_i, sty_i)
            numfmts <- c(numfmts, self$styles_mgr$get_numfmt_id(sty_i))
          } else {
            numfmts <- c(numfmts, as_xml_attr(numfmt[[i]]))
          }
        }
      }

      if (is.null(params$name) && !is.null(pivot_table))
        params$name <- pivot_table

      # not sure if rows & cols can be formulas too
      if (any(sel <- !data %in% names(x))) {

        varfun <- data[sel]

        if (is.null(names(varfun))) {
          stop("Unknown variable in data argument: ", data)
        }

        for (var in varfun) {
          if (all(grepl("^=", names(varfun)))) {
            x[[var]] <- names(varfun[varfun == var])
            class(x[[var]]) <- c("is_formula", "character")
          } else {
            stop("missing variable found in pivot table: data object. Formula names must begin with '='.")
          }
        }

      }

      pivot_table <- create_pivot_table(
        x       = x,
        dims    = dims,
        filter  = filter,
        rows    = rows,
        cols    = cols,
        data    = data,
        n       = length(self$pivotTables) + 1L,
        fun     = fun,
        params  = params,
        numfmts = numfmts
      )

      if (missing(filter))   filter   <- ""
      if (missing(rows))     rows     <- ""
      if (missing(cols))     cols     <- ""
      if (missing(data))     data     <- ""
      if (missing(slicer))   slicer   <- ""
      if (missing(timeline)) timeline <- ""

      self$append("pivotTables", pivot_table)
      cacheId <- length(self$pivotTables)
      self$worksheets[[sheet]]$relships$pivotTable <- append(
        self$worksheets[[sheet]]$relships$pivotTable,
        cacheId
      )

      self$append("pivotDefinitions", pivot_def_xml(x, filter, rows, cols, data, slicer, timeline, cacheId))

      self$append("pivotDefinitionsRels", pivot_def_rel(cacheId))
      self$append("pivotRecords", pivot_rec_xml(x))
      self$append("pivotTables.xml.rels", pivot_xml_rels(cacheId))


      rId <- paste0("rId", 20000 + cacheId)

      pivotCache <- sprintf(
        "<pivotCache cacheId=\"%s\" r:id=\"%s\"/>",
        cacheId,
        rId
      )
      if (length(self$workbook$pivotCaches)) {
        self$workbook$pivotCaches <- xml_add_child(self$workbook$pivotCaches, xml_child = pivotCache)
      } else {
        self$workbook$pivotCaches <- xml_node_create("pivotCaches", xml_children = pivotCache)
      }

      next_id <- get_next_id(self$worksheets_rels[[sheet]])

      self$worksheets_rels[[sheet]] <- c(
        self$worksheets_rels[[sheet]],
        sprintf(
          '<Relationship Id=\"%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable%s.xml\"/>',
          next_id,
          cacheId
        )
      )

      self$append("workbook.xml.rels",
        sprintf(
          "<Relationship Id=\"%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition%s.xml\"/>",
          rId,
          cacheId
        )
      )

      self$append("Content_Types",
                  c(
                    sprintf("<Override PartName=\"/xl/pivotTables/pivotTable%s.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"/>", cacheId),
                    sprintf("<Override PartName=\"/xl/pivotCache/pivotCacheDefinition%s.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/>", cacheId),
                    sprintf("<Override PartName=\"/xl/pivotCache/pivotCacheRecords%s.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml\"/>", cacheId)
                  )
      )

      self$worksheets[[sheet]]$sheetFormatPr <- "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\" x14ac:dyDescent=\"0.2\"/>"

      invisible(self)
    },

    #' @description add pivot table
    #' @param x a wb_data object
    #' @param dims the worksheet cell where the pivot table is placed
    #' @param pivot_table the name of a pivot table on the selected sheet
    #' @param slicer a variable used as slicer for the pivot table
    #' @param params a list of parameters to modify pivot table creation
    #' @return The `wbWorkbook` object
    add_slicer = function(x, dims = "A1", sheet = current_sheet(), pivot_table, slicer, params) {

      if (!grepl(":", dims)) {
        ddims <- dims_to_rowcol(dims, as_integer = TRUE)

        dims <- rowcol_to_dims(
          row = c(ddims[["row"]], ddims[["row"]] + 12L),
          col = c(ddims[["col"]], ddims[["col"]] + 1L)
        )
      }

      if (missing(x))
        stop("x cannot be missing in add_slicer")

      assert_class(x, "wb_data")
      if (missing(params)) {
        params <- NULL
      } else {
        arguments <- c(
          "caption", "choose", "column_count", "cross_filter", "edit_as",
          "hide_no_data_items", "level", "locked_position", "row_height",
          "show_caption", "show_missing", "sort_order", "start_item", "style"
        )
        params <- standardize_case_names(params, arguments = arguments, return = TRUE)
      }

      sheet <- private$get_sheet_index(sheet)

      pt <- rbindlist(xml_attr(self$pivotTables, "pivotTableDefinition"))
      sel <- which(pt$name == pivot_table)
      cid <- pt$cacheId[sel]

      uni_name <- paste0(stringi::stri_replace_all_fixed(slicer, " ", "_"), cid)

      ### slicer_cache
      sortOrder <- NULL
      if (!is.null(params$sort_order))
        sortOrder <- params$sort_order

      showMissing <- NULL
      if (!is.null(params$show_missing))
        showMissing <- params$show_missing

      crossFilter <- NULL
      if (!is.null(params$cross_filter))
        crossFilter <- params$cross_filter

      # TODO we might be able to initialize the field from here. Something like
      # get_item(...) and insert it to the pivotDefinition

      # test that slicer is initalized in wb$pivotDefinitions.
      pt  <- self$worksheets[[sheet]]$relships$pivotTable
      ptl <- rbindlist(xml_attr(self$pivotTables[pt], "pivotTableDefinition"))
      pt  <- pt[which(ptl$name == pivot_table)]

      fields <- xml_node(self$pivotDefinitions[pt], "pivotCacheDefinition", "cacheFields", "cacheField")
      names(fields) <- vapply(xml_attr(fields, "cacheField"), function(x) x[["name"]], "")

      if (is.na(xml_attr(fields[slicer], "cacheField", "sharedItems")[[1]]["count"])) {
        stop("slicer was not initialized in pivot table!")
      }

      choose    <- params$choose

      if (!is.null(choose) && !is.na(choose[slicer])) {
        choo <- choose[slicer]
      } else {
        choo <- NULL
      }

      tab_xml <- xml_node_create(
        "tabular",
        xml_attributes = c(
          pivotCacheId = cid,
          sortOrder    = sortOrder,
          showMissing  = showMissing,
          crossFilter  = crossFilter
        ),
        xml_children = get_items(x, which(names(x) == slicer), NULL, slicer = TRUE, choose = choo, has_default = TRUE)
      )


      hide_items_with_no_data <- ""
      if (isTRUE(params$hide_no_data_items)) {
        hide_items_with_no_data <- '<extLst>
          <x:ext xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" uri="{470722E0-AACD-4C17-9CDC-17EF765DBC7E}">
            <x15:slicerCacheHideItemsWithNoData/>
          </x:ext>
        </extLst>'
      }

      slicer_cache <- read_xml(sprintf(
        '<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" name="Slicer_%s" xr10:uid="{72B411E0-23B7-7444-B533-EAC1856BE56A}" sourceName="%s">
          <pivotTables>
            <pivotTable tabId="%s" name="%s" />
          </pivotTables>
          <data>
            %s
          </data>
          %s
        </slicerCacheDefinition>',
        uni_name,
        slicer,
        sheet,
        pivot_table,
        tab_xml,
        hide_items_with_no_data
      ), pointer = FALSE)

      # we need the slicer cache
      self$append(
        "slicerCaches",
        slicer_cache
      )

      # and the actual slicer
      if (length(self$worksheets[[sheet]]$relships$slicer) == 0) {
        self$append(
          "slicers",
          '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"></slicers>'
        )
        self$worksheets[[sheet]]$relships$slicer <- length(self$slicers)
      }

      caption <- slicer
      if (!is.null(params$caption))
        caption <- params$caption

      row_height <- 230716
      if (!is.null(params$row_height))
        row_height <- params$row_height

      column_count <- NULL
      if (!is.null(params$column_count))
        column_count <- params$column_count

      style <- NULL
      if (!is.null(params$style))
        style <- params$style

      startItem <- NULL
      if (!is.null(params$start_item))
        startItem <- params$start_item

      showCaption <- NULL
      if (!is.null(params$show_caption))
        showCaption <- params$show_caption

      level <- NULL
      if (!is.null(params$level))
        level <- params$level

      lockedPosition <- NULL
      if (!is.null(params$locked_position))
        lockedPosition <- params$locked_position

      slicer_xml <- xml_node_create(
        "slicer",
        xml_attributes = c(
          name           = uni_name,
          `xr10:uid`     = st_guid(),
          cache          = paste0("Slicer_", uni_name),
          caption        = caption,
          rowHeight      = as_xml_attr(row_height),
          columnCount    = as_xml_attr(column_count),
          startItem      = as_xml_attr(startItem),
          showCaption    = as_xml_attr(showCaption),
          level          = as_xml_attr(level),
          lockedPosition = as_xml_attr(lockedPosition),
          style          = style
        )
      )

      sel <- self$worksheets[[sheet]]$relships$slicer
      self$slicers[sel] <- xml_add_child(self$slicers[sel], xml_child = slicer_xml)

      slicer_id <- length(self$slicerCaches)
      # append it to the workbook.xml.rels
      self$append(
        "workbook.xml.rels",
        sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicerCache\" Target=\"slicerCaches/slicerCache%s.xml\"/>",
                100000 + slicer_id, slicer_id)
      )

      # add this defined name
      self$workbook$definedNames <- append(
        self$workbook$definedNames,
        sprintf("<definedName name=\"Slicer_%s\">#N/A</definedName>", uni_name)
      )

      # add the workbook extension list
      if (is.null(self$workbook$extLst)) {
        self$workbook$extLst <- xml_node_create("extLst")
      }

      if (!grepl("xmlns:x14", self$workbook$extLst)) {
        self$workbook$extLst <- xml_add_child(self$workbook$extLst, xml_child = '<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"></ext>')
      }

      extLst_xml <- xml_node(self$workbook$extLst, "extLst", "ext")
      is_ext_x14 <- grepl("xmlns:x14", extLst_xml)

      # check if node has x14:slicerCaches
      if (!grepl("<x14:slicerCaches>", extLst_xml[is_ext_x14])) {
        extLst_xml[is_ext_x14] <- xml_add_child(
          extLst_xml[is_ext_x14],
          xml_child = xml_node_create("x14:slicerCaches")
        )
      }

      extLst_xml[is_ext_x14] <- xml_add_child(
        extLst_xml[is_ext_x14],
        xml_child = sprintf('<x14:slicerCache r:id="rId%s" />', 100000 + slicer_id),
        level = "x14:slicerCaches"
      )

      self$workbook$extLst <- xml_node_create("extLst", xml_children = extLst_xml)

      # add a drawing for the slicer
      drawing_xml <- read_xml(sprintf('
        <xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">
          <xdr:absoluteAnchor>
            <xdr:pos x="0" y="0" />
            <xdr:ext cx="9313333" cy="6070985" />
            <mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">
              <mc:Choice xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\">
                <xdr:graphicFrame macro=\"\">
                  <xdr:nvGraphicFramePr>
                    <xdr:cNvPr id=\"2\" name=\"%s\">
                      <a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{54EAEA0F-6B31-AC4D-D672-2AA8C6402920}\"/></a:ext></a:extLst>
                    </xdr:cNvPr>
                    <xdr:cNvGraphicFramePr/>
                  </xdr:nvGraphicFramePr>
                  <xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>
                  <a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" name=\"%s\"/></a:graphicData></a:graphic>
                </xdr:graphicFrame>
              </mc:Choice>
              <mc:Fallback><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"6959600\" y=\"2794000\"/><a:ext cx=\"1828800\" cy=\"2428869\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-GB\" sz=\"1100\"/><a:t>This shape represents a slicer. Slicers are supported in Excel 2010 or later.\n\nIf the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2003 or earlier, the slicer cannot be used.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>
            </mc:AlternateContent>
            <xdr:clientData/>
          </xdr:absoluteAnchor>
        </xdr:wsDr>
        ', uni_name, uni_name
      ), pointer = FALSE)


      edit_as <- "oneCell"
      if (!is.null(params$edit_as))
        edit_as <- params$edit_as

      # place the drawing
      self$add_drawing(dims = dims, sheet = sheet, xml = drawing_xml, edit_as = edit_as)


      next_id <- get_next_id(self$worksheets_rels[[sheet]])



      # add the pivot table and the drawing to the worksheet
      if (!any(grepl(sprintf("Target=\"../slicers/slicer%s.xml\"", self$worksheets[[sheet]]$relships$slicer), self$worksheets_rels[[sheet]]))) {

        slicer_list_xml <- sprintf(
          '<x14:slicer r:id=\"%s\"/>',
          next_id
        )

        self$worksheets[[sheet]]$.__enclos_env__$private$do_append_x14(slicer_list_xml, "x14:slicer", "x14:slicerList")

        self$worksheets_rels[[sheet]] <- append(
          self$worksheets_rels[[sheet]],
          sprintf(
            "<Relationship Id=\"%s\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicer\" Target=\"../slicers/slicer%s.xml\"/>",
            next_id, self$worksheets[[sheet]]$relships$slicer
          )
        )

      }

      slicer_xml <- sprintf(
        "<Override PartName=\"/xl/slicers/slicer%s.xml\" ContentType=\"application/vnd.ms-excel.slicer+xml\"/>",
        self$worksheets[[sheet]]$relships$slicer
      )

      if (!any(self$Content_Types == slicer_xml)) {
        self$append(
          "Content_Types",
          slicer_xml
        )
      }

      value <- sprintf(
        "<Override PartName=\"/xl/slicerCaches/slicerCache%s.xml\" ContentType=\"application/vnd.ms-excel.slicerCache+xml\"/>",
        slicer_id
      )

      self$append(
        "Content_Types", value
      )

      invisible(self)
    },

    #' @description add pivot table
    #' @return The `wbWorkbook` object
    remove_slicer = function(sheet = current_sheet()) {
      sheet <- private$get_sheet_index(sheet)

      # get indices
      slicer_id       <- self$worksheets[[sheet]]$relships$slicer

      # skip if nothing to do
      if (identical(slicer_id, integer())) return(invisible(self))

      cache_names     <- unname(sapply(xml_attr(self$slicers[slicer_id], "slicers", "slicer"), "[", "cache"))
      slicer_names    <- unname(sapply(xml_attr(self$slicerCaches, "slicerCacheDefinition"), "[", "name"))
      slicer_cache_id <- which(cache_names %in% slicer_names)

      # strings to grep
      slicer_xml <- sprintf("slicers/slicer%s.xml", slicer_id)
      caches_xml <- sprintf("slicerCaches/slicerCache%s.xml", slicer_cache_id)

      # empty slicer
      self$slicers[slicer_id]                  <- ""
      # empty slicerCache
      self$slicerCaches[slicer_cache_id]       <- ""

      # remove slicer cache relship
      self$worksheets[[sheet]]$relships$slicer <- integer()
      # remove worksheet relationship
      self$worksheets_rels[[sheet]]            <- self$worksheets_rels[[sheet]][!grepl(slicer_xml, self$worksheets_rels[[sheet]])]
      # remove "x14:slicerList"
      is_ext_x14 <- grepl("x14:slicerList", self$worksheets[[sheet]]$extLst)
      extLst     <- xml_rm_child(self$worksheets[[sheet]]$extLst[is_ext_x14], xml_child = "x14:slicerList")
      self$worksheets[[sheet]]$extLst[is_ext_x14] <- extLst

      # clear workbook.xml.rels
      self$workbook.xml.rels                   <- self$workbook.xml.rels[!grepl(paste0(caches_xml, collapse = "|"), self$workbook.xml.rels)]

      # clear Content_Types
      self$Content_Types                       <- self$Content_Types[!grepl(paste0(c(slicer_xml, caches_xml), collapse = "|"), self$Content_Types)]

      invisible(self)
    },

    #' @description add pivot table
    #' @param x a wb_data object
    #' @param dims the worksheet cell where the pivot table is placed
    #' @param pivot_table the name of a pivot table on the selected sheet
    #' @param timeline a variable used as timeline for the pivot table
    #' @param params a list of parameters to modify pivot table creation
    #' @return The `wbWorkbook` object
    add_timeline = function(x, dims = "A1", sheet = current_sheet(), pivot_table, timeline, params) {

      if (!grepl(":", dims)) {
        ddims <- dims_to_rowcol(dims, as_integer = TRUE)

        dims <- rowcol_to_dims(
          row = c(ddims[["row"]], ddims[["row"]] + 12L),
          col = c(ddims[["col"]], ddims[["col"]] + 1L)
        )
      }

      if (missing(x))
        stop("x cannot be missing in add_timeline")

      assert_class(x, "wb_data")
      if (missing(params)) {
        params <- NULL
      } else {
        arguments <- c(
          "beg_date", "caption", "choose_beg", "choose_end", "column_count",
          "cross_filter", "edit_as", "end_date", "level", "row_height",
          "scroll_position", "selection_level", "show_header",
          "show_horizontal_scrollbar", "show_missing", "show_selection_label",
          "show_time_level", "sort_order", "style"
        )
        params <- standardize_case_names(params, arguments = arguments, return = TRUE)
      }

      sheet <- private$get_sheet_index(sheet)

      pt <- rbindlist(xml_attr(self$pivotTables, "pivotTableDefinition"))
      sel <- which(pt$name == pivot_table)
      cid <- pt$cacheId[sel]

      uni_name <- paste0(stringi::stri_replace_all_fixed(timeline, " ", "_"), cid)

      # TODO we might be able to initialize the field from here. Something like
      # get_item(...) and insert it to the pivotDefinition

      # test that slicer is initalized in wb$pivotDefinitions.
      pt  <- self$worksheets[[sheet]]$relships$pivotTable
      ptl <- rbindlist(xml_attr(self$pivotTables[pt], "pivotTableDefinition"))
      pt  <- pt[which(ptl$name == pivot_table)]

      fields <- xml_node(self$pivotDefinitions[pt], "pivotCacheDefinition", "cacheFields", "cacheField")
      names(fields) <- vapply(xml_attr(fields, "cacheField"), function(x) x[["name"]], "")

      if (is.na(xml_attr(fields[timeline], "cacheField", "sharedItems")[[1]]["count"])) {
        stop("timeline was not initialized in pivot table!")
      }

      if (!inherits(x[[timeline]], "Date") && !inherits(x[[timeline]], "POSIXt")) {
        stop("a timeline must be a date or a POSIXt object")
      } else {
        startDate <- min(x[[timeline]],  na.rm = TRUE)
        endDate   <- max(x[[timeline]],  na.rm = TRUE)
        meanDate  <- mean(x[[timeline]], na.rm = TRUE)
      }

      beg_date <- params$beg_date
      if (!is.null(beg_date)) {
        startDate <- beg_date
      }

      end_date <- params$end_date
      if (!is.null(end_date)) {
        endDate <- end_date
      }

      choose_beg <- params$choose_beg
      if (is.null(choose_beg)) {
        choose_beg <- startDate
      }
      if (!inherits(choose_beg, "Date") && !inherits(choose_beg, "POSIXt")) {
        stop("choose_beg: a timeline must be a date or a POSIXt object")
      }

      choose_end <- params$choose_end
      if (is.null(choose_end)) {
        choose_end <- endDate
      }
      if (!inherits(choose_end, "Date") && !inherits(choose_end, "POSIXt")) {
        stop("choose_end: a timeline must be a date or a POSIXt object")
      }


      selection_xml <- xml_node_create(
        "selection",
        xml_attributes = c(
          startDate = format(as_POSIXct_utc(choose_beg), "%Y-%m-%dT%H:%M:%SZ"),
          endDate   = format(as_POSIXct_utc(choose_end), "%Y-%m-%dT%H:%M:%SZ")
        )
      )

      bounds_xml <- xml_node_create(
        "bounds",
        xml_attributes = c(
          startDate = format(as_POSIXct_utc(startDate), "%Y-%m-%dT%H:%M:%SZ"),
          endDate   = format(as_POSIXct_utc(endDate),   "%Y-%m-%dT%H:%M:%SZ")
        )
      )

      # without choose: filterType = unknown
      timeline_cache <- read_xml(sprintf(
        '<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" mc:Ignorable="xr10" name="NativeTimeline_%s" xr10:uid="%s" sourceName="%s">
          <pivotTables>
            <pivotTable tabId="%s" name="%s" />
          </pivotTables>
          <state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="%s" filterType="dateBetween">
            %s
            %s
          </state>
        </timelineCacheDefinition>',
        uni_name,
        st_guid(),
        timeline,
        sheet,
        pivot_table,
        cid,
        selection_xml,
        bounds_xml
      ), pointer = FALSE)

      # we need the timeline cache
      self$append(
        "timelineCaches",
        timeline_cache
      )

      # and the actual slicer
      if (length(self$worksheets[[sheet]]$relships$timeline) == 0) {
        self$append(
          "timelines",
          '<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" mc:Ignorable="x xr10"></timelines>'
        )
        self$worksheets[[sheet]]$relships$timeline <- length(self$timelines)
      }

      caption <- timeline
      if (!is.null(params$caption))
        caption <- params$caption

      level <- 2
      if (!is.null(params$level))
        level <- params$level

      selection_level <- 2
      if (!is.null(params$selection_level))
        selection_level <- params$selection_level

      scroll_position <- meanDate
      if (!is.null(params$scroll_position))
        scroll_position <- params$scroll_position

      style <- NULL
      if (!is.null(params$style))
        style <- params$style

      showHeader <- NULL
      if (!is.null(params$show_header))
        showHeader <- params$show_header

      showSelectionLabel <- NULL
      if (!is.null(params$show_selection_label))
        showSelectionLabel <- params$show_selection_label

      showTimeLevel <- NULL
      if (!is.null(params$show_time_level))
        showTimeLevel <- params$show_time_level

      showHorizontalScrollbar <- NULL
      if (!is.null(params$show_horizontal_scrollbar))
        showHorizontalScrollbar <- params$show_horizontal_scrollbar

      timeline_xml <- xml_node_create(
        "timeline",
        xml_attributes = c(
          name                    = uni_name,
          `xr10:uid`              = st_guid(),
          cache                   = paste0("NativeTimeline_", uni_name),
          caption                 = caption,
          level                   = as_xml_attr(level),
          selectionLevel          = as_xml_attr(selection_level),
          scrollPosition          = format(as_POSIXct_utc(scroll_position), "%Y-%m-%dT%H:%M:%SZ"),
          showHeader              = as_xml_attr(showHeader),
          showSelectionLabel      = as_xml_attr(showSelectionLabel),
          showTimeLevel           = as_xml_attr(showTimeLevel),
          showHorizontalScrollbar = as_xml_attr(showHorizontalScrollbar),
          style                   = style
        )
      )

      sel <- self$worksheets[[sheet]]$relships$timeline
      self$timelines[sel] <- xml_add_child(self$timelines[sel], xml_child = timeline_xml)

      timeline_id <- length(self$timelineCaches)
      # append it to the workbook.xml.rels
      self$append(
        "workbook.xml.rels",
        sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.microsoft.com/office/2011/relationships/timelineCache\" Target=\"timelineCaches/timelineCache%s.xml\"/>",
                200000 + timeline_id, timeline_id)
      )

      # add this defined name
      self$workbook$definedNames <- append(
        self$workbook$definedNames,
        sprintf("<definedName name=\"NativeTimeline_%s\">#N/A</definedName>", uni_name)
      )

      # add the workbook extension list
      if (is.null(self$workbook$extLst)) {
        self$workbook$extLst <- xml_node_create("extLst")
      }

      if (!grepl("xmlns:x15", self$workbook$extLst)) {
        self$workbook$extLst <- xml_add_child(self$workbook$extLst, xml_child = '<ext xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}"></ext>')
      }

      extLst_xml <- xml_node(self$workbook$extLst, "extLst", "ext")
      is_ext_x15 <- grepl("xmlns:x15", extLst_xml)

      # check if node has x15:timelineCacheRefs
      if (!grepl("<x15:timelineCacheRefs>", extLst_xml[is_ext_x15])) {
        extLst_xml[is_ext_x15] <- xml_add_child(
          extLst_xml[is_ext_x15],
          xml_child = xml_node_create("x15:timelineCacheRefs")
        )
      }

      extLst_xml[is_ext_x15] <- xml_add_child(
        extLst_xml[is_ext_x15],
        xml_child = sprintf('<x15:timelineCacheRef r:id="rId%s" />', 200000 + timeline_id),
        level = "x15:timelineCacheRefs"
      )

      self$workbook$extLst <- xml_node_create("extLst", xml_children = extLst_xml)

      # add a drawing for the slicer
      drawing_xml <- read_xml(sprintf('
        <xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">
          <xdr:absoluteAnchor>
            <xdr:pos x="0" y="0" />
            <xdr:ext cx="9313333" cy="6070985" />
            <mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">
              <mc:Choice xmlns:tsle=\"http://schemas.microsoft.com/office/drawing/2012/timeslicer\" Requires=\"tsle\">
                <xdr:graphicFrame macro=\"\">
                  <xdr:nvGraphicFramePr>
                    <xdr:cNvPr id=\"2\" name="%s">
                      <a:extLst><a:ext uri=\"{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F0ED9F26-D0CD-D3F1-9FDD-EABC216430BF}\"/></a:ext></a:extLst>
                    </xdr:cNvPr>
                    <xdr:cNvGraphicFramePr/>
                  </xdr:nvGraphicFramePr>
                  <xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>
                  <a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2012/timeslicer\"><tsle:timeslicer xmlns:tsle=\"http://schemas.microsoft.com/office/drawing/2012/timeslicer\" name=\"%s\"/></a:graphicData></a:graphic>
                </xdr:graphicFrame>
              </mc:Choice>
              <mc:Fallback><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"6019800\" y=\"3632200\"/><a:ext cx=\"3340100\" cy=\"1320800\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-GB\" sz=\"1100\"/><a:t>Timeline: Works in Excel 2013 or higher. Do not move or resize.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>
            </mc:AlternateContent>
            <xdr:clientData/>
          </xdr:absoluteAnchor>
        </xdr:wsDr>
        ', uni_name, uni_name
      ), pointer = FALSE)


      edit_as <- "oneCell"
      if (!is.null(params$edit_as))
        edit_as <- params$edit_as

      # place the drawing
      self$add_drawing(dims = dims, sheet = sheet, xml = drawing_xml, edit_as = edit_as)


      next_id <- get_next_id(self$worksheets_rels[[sheet]])

      # add the pivot table and the drawing to the worksheet
      if (!any(grepl(sprintf("Target=\"../timelines/timeline%s.xml\"", self$worksheets[[sheet]]$relships$timeline), self$worksheets_rels[[sheet]]))) {

        timeline_list_xml <- sprintf(
          '<x15:timelineRefs><x15:timelineRef r:id=\"%s\"/></x15:timelineRefs>',
          next_id
        )

        # add the extension list to the worksheet
        is_ext_x15 <- grepl("x15:timelineRefs", self$worksheets[[sheet]]$extLst)
        if (length(self$worksheets[[sheet]]$extLst) == 0 || !any(is_ext_x15)) {

          ext_x15 <- "<ext uri=\"{7E03D99C-DC04-49d9-9315-930204A7B6E9}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"></ext>"

          self$worksheets[[sheet]]$append(
            "extLst",
            ext_x15
          )

          is_ext_x15 <- length(self$worksheets[[sheet]]$extLst)

        }

        self$worksheets[[sheet]]$extLst[is_ext_x15] <- xml_add_child(
          self$worksheets[[sheet]]$extLst[is_ext_x15],
          xml_child = timeline_list_xml
        )

        self$worksheets_rels[[sheet]] <- append(
          self$worksheets_rels[[sheet]],
          sprintf(
            "<Relationship Id=\"%s\" Type=\"http://schemas.microsoft.com/office/2011/relationships/timeline\" Target=\"../timelines/timeline%s.xml\"/>",
            next_id, self$worksheets[[sheet]]$relships$timeline
          )
        )

      }

      timeline_xml <- sprintf(
        "<Override PartName=\"/xl/timelines/timeline%s.xml\" ContentType=\"application/vnd.ms-excel.timeline+xml\"/>",
        self$worksheets[[sheet]]$relships$timeline
      )

      if (!any(self$Content_Types == timeline_xml)) {
        self$append(
          "Content_Types",
          timeline_xml
        )
      }

      value <- sprintf(
        "<Override PartName=\"/xl/timelineCaches/timelineCache%s.xml\" ContentType=\"application/vnd.ms-excel.timelineCache+xml\"/>",
        timeline_id
      )

      self$append(
        "Content_Types", value
      )

      invisible(self)
    },

    #' @description add pivot table
    #' @return The `wbWorkbook` object
    remove_timeline = function(sheet = current_sheet()) {
      sheet <- private$get_sheet_index(sheet)

      # get indices
      timeline_id       <- self$worksheets[[sheet]]$relships$timeline

      # skip if nothing to do
      if (identical(timeline_id, integer())) return(invisible(self))

      cache_names       <- unname(sapply(xml_attr(self$timelines[timeline_id], "timelines", "timeline"), "[", "cache"))
      timeline_names    <- unname(sapply(xml_attr(self$timelineCaches, "timelineCacheDefinition"), "[", "name"))
      timeline_cache_id <- which(cache_names %in% timeline_names)

      # strings to grep
      timeline_xml <- sprintf("timelines/timeline%s.xml", timeline_id)
      caches_xml   <- sprintf("timelineCaches/timelineCache%s.xml", timeline_cache_id)

      # empty timelines
      self$timelines[timeline_id]                <- ""
      # empty timelineCache
      self$timelineCaches[timeline_cache_id]     <- ""

      # remove timeline cache relship
      self$worksheets[[sheet]]$relships$timeline <- integer()
      # remove worksheet relationship
      self$worksheets_rels[[sheet]]              <- self$worksheets_rels[[sheet]][!grepl(timeline_xml, self$worksheets_rels[[sheet]])]
      # remove "x15:timelineRefs"
      is_ext_x15 <- grepl("x15:timelineRefs", self$worksheets[[sheet]]$extLst)
      extLst     <- xml_rm_child(self$worksheets[[sheet]]$extLst[is_ext_x15], xml_child = "x15:timelineRefs")
      self$worksheets[[sheet]]$extLst[is_ext_x15] <- extLst

      # clear workbook.xml.rels
      self$workbook.xml.rels                     <- self$workbook.xml.rels[!grepl(paste0(caches_xml, collapse = "|"), self$workbook.xml.rels)]

      # clear Content_Types
      self$Content_Types                         <- self$Content_Types[!grepl(paste0(c(timeline_xml, caches_xml), collapse = "|"), self$Content_Types)]

      invisible(self)
    },

    #' @description Add formula
    #' @param x x
    #' @param start_col startCol
    #' @param start_row startRow
    #' @param array array
    #' @param cm cm
    #' @param apply_cell_style applyCellStyle
    #' @param remove_cell_style if writing into existing cells, should the cell style be removed?
    #' @param enforce enforce dims
    #' @param shared shared formula
    #' @param name name
    #' @return The `wbWorkbook` object
    add_formula = function(
        sheet             = current_sheet(),
        x,
        dims              = wb_dims(start_row, start_col),
        start_col         = 1,
        start_row         = 1,
        array             = FALSE,
        cm                = FALSE,
        apply_cell_style  = TRUE,
        remove_cell_style = FALSE,
        enforce           = FALSE,
        shared            = FALSE,
        name              = NULL,
        ...
    ) {

      standardize_case_names(...)

      if (is.character(x) && !is.null(names(x)) && is.null(name)) {
        assert_class(x, "character")
        assert_named_region(names(x))

        if (NROW(nr <- self$get_named_regions())) {
          nr_name <- nr$name[nr$local == 0]

          if (any(tolower(names(x)) %in% tolower(nr_name)))
            stop("named regions cannot be duplicates")
        }

        xml <- xml_node_create(
          "definedName",
          xml_children = x,
          xml_attributes = c(name = names(x))
        )
        private$append_workbook_field("definedNames", xml)

        message("formula registered to the workbook")
        return(invisible(self))
      }

      do_write_formula(
        wb                = self,
        sheet             = sheet,
        x                 = x,
        start_col         = start_col,
        start_row         = start_row,
        dims              = dims,
        array             = array,
        cm                = cm,
        apply_cell_style  = apply_cell_style,
        remove_cell_style = remove_cell_style,
        enforce           = enforce,
        shared            = shared,
        name              = name
      )
      invisible(self)
    },

    #' @description Add hyperlink
    #' @param sheet sheet
    #' @param dims dims
    #' @param target target
    #' @param tooltip tooltip
    #' @param is_external is_external
    #' @param col_names col_names
    #' @return The `wbWorkbook` object
    add_hyperlink = function(
      sheet       = current_sheet(),
      dims        = "A1",
      target      = NULL,
      tooltip     = NULL,
      is_external = TRUE,
      col_names   = FALSE
    ) {

      sheet <- private$get_sheet_index(sheet)

      if (!grepl(":", dims)) col_names <- FALSE

      x <- wb_to_df(self, sheet = sheet, dims = dims, col_names = col_names)
      nams <- names(x)

      if (!is.null(target) && is.null(names(target))) {
        if (nrow(x) > ncol(x)) {
          target <- as.data.frame(as.matrix(target, nrow = nrow(x), ncol = ncol(x)), stringsAsFactors = FALSE)
        }
        names(target) <- nams
      }

      if (!is.null(tooltip) && is.null(names(tooltip))) {
        if (nrow(x) > ncol(x)) {
          tooltip <- as.data.frame(as.matrix(tooltip, nrow = nrow(x), ncol = ncol(x)), stringsAsFactors = FALSE)
        }
        names(tooltip) <- nams
      }

      rel_ids <- NULL
      if (length(self$worksheets_rels[[sheet]])) {
        relships <- rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))
        rel_ids  <- as.integer(gsub("\\D+", "", relships$Id))
      }

      max_id <- max(rel_ids, 0)
      if (!is.null(nams) && !all(nams %in% names(x)))
        stop("some selected columns are not part of `dims`")

      if (!is.null(target) && !any(names(target) %in% nams)) {
        warning("target not found in selected `dims`")
        return(invisible(self))
      }

      if (!is.null(tooltip) && !any(names(tooltip) %in% nams)) {
        warning("tooltip not found in selected `dims`")
        return(invisible(self))
      }

      for (nam in nams) {

        ddims <- dims_to_dataframe(dims, fill = TRUE)
        names(ddims) <- nams
        # the first row is removed, because it is used only
        # to identify the column, if a target/tooltip is named
        if (col_names) ddims <- ddims[-1, , drop = FALSE]

        if (!is.null(tooltip) && nam %in% names(tooltip)) {
          tooltip_i <- tooltip[[nam]]
        } else {
          tooltip_i <- NULL
        }

        if (is_external) {

          Target     <- x[[nam]]
          if (!is.null(target) && nam %in% names(target)) {
            Target <- target[[nam]]
          }
          max_id_seq <- seq.int(from = max_id + 1L, length.out = length(Target))
          Id         <- paste0("rId", max_id_seq)
          TargetMode <- "External"

          # display <- target
          df <- data.frame(
            Id = Id,
            Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            Target = Target,
            TargetMode = TargetMode,
            stringsAsFactors = FALSE
          )

          new_relship <- df_to_xml("Relationship", df)

          self$worksheets_rels[[sheet]] <- append(
            self$worksheets_rels[[sheet]],
            new_relship
          )

          df <- data.frame(
            ref = unlist(unname(ddims[[nam]])),
            `r:id` = Id,
            tooltip = as_xml_attr(tooltip_i),
            stringsAsFactors = FALSE,
            check.names = FALSE
          )

          max_id <- max(max_id_seq, 0) + 1L

        } else { # a cell reference within the workbook

          if (!is.null(target)) {
            Location <- unname(target[[nam]])
            Display  <- unname(x[[nam]])
          } else {
            Location <- unname(x[[nam]])
            Display  <- as_xml_attr(NULL)
          }

          df <- data.frame(
            ref = unlist(unname(ddims[[nam]])),
            location = Location,
            display = Display,
            tooltip = as_xml_attr(tooltip_i),
            stringsAsFactors = FALSE,
            check.names = FALSE
          )

        }

        new_hyperlink <- df_to_xml("hyperlink", df)

        self$worksheets[[sheet]]$hyperlinks <- append(
          unlist(self$worksheets[[sheet]]$hyperlinks),
          new_hyperlink
        )

        # get hyperlink color from template
        if (is.null(self$theme)) {
          has_hlink <- 11
        } else {
          clrs <- xml_node(self$theme, "a:theme", "a:themeElements", "a:clrScheme")
          has_hlink <- which(xml_node_name(clrs, "a:clrScheme") == "a:hlink")
        }

        if (has_hlink) {
          hyperlink_col <- wb_color(theme = has_hlink - 1L)
        } else {
          hyperlink_col <- wb_color(hex = "FF0000FF")
        }

        self$add_font(
          sheet     = sheet,
          dims      = ddims[[nam]],
          color     = hyperlink_col,
          name      = self$get_base_font()$name$val,
          size      = self$get_base_font()$size$val,
          underline = "single"
        )

      } # end nam loop

      if (length(self$worksheets_rels[[sheet]])) {
        relships <- rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))
        rel_ids  <- as.integer(gsub("\\D+", "", relships$Id[basename(relships$Type) == "hyperlink"]))
        self$worksheets[[sheet]]$relships$hyperlink <- rel_ids
      }

      invisible(self)
    },

    #' @description remove hyperlink
    #' @param sheet sheet
    #' @param dims dims
    #' @return The `wbWorkbook` object
    remove_hyperlink = function(sheet = current_sheet(), dims = NULL) {

      sheet <- private$get_sheet_index(sheet)

      # get all hyperlinks
      hls    <- self$worksheets[[sheet]]$hyperlinks

      if (length(hls)) {
        hls_df <- rbindlist(xml_attr(hls, "hyperlink"))

        if (is.null(dims)) {
          # remove all hyperlinks
          self$worksheets[[sheet]]$hyperlinks <- character()
          refs <- hls_df$ref
        } else {
          # get cells in dims, get required cells, replace these and reduce refs
          ddims <- dims_to_dataframe(dims = dims, fill = TRUE)
          sel <- which(hls_df$ref %in% unname(unlist(ddims)))
          self$worksheets[[sheet]]$hyperlinks <- hls_df$ref[-sel]
          refs <- hls_df$ref[sel]
        }

        # TODO remove "r:id" reference from worksheets_rels

        # reset font style
        for (ref in refs) {
          self$add_cell_style(font_id = 0)
        }
      }

      invisible(self)
    },

    #' @description add style
    #' @param style style
    #' @param style_name style_name
    #' @return The `wbWorkbook` object
    add_style = function(style = NULL, style_name = NULL) {

      assert_class(style, "character")

      if (is.null(style_name)) {
        style_name <- deparse(substitute(style))

        if (xml_node_name(style) == "tableStyle")
          style_name <- rbindlist(xml_attr(style, "tableStyle"))$name

      } else {
        assert_class(style_name, "character")
      }

      self$styles_mgr$add(style, style_name)

      invisible(self)
    },

    ### to dataframe ----
    #' @description to_df
    #' @param sheet Either sheet name or index. When missing the first sheet in the workbook is selected.
    #' @param col_names If TRUE, the first row of data will be used as column names.
    #' @param row_names If TRUE, the first col of data will be used as row names.
    #' @param dims Character string of type "A1:B2" as optional dimensions to be imported.
    #' @param detect_dates If TRUE, attempt to recognize dates and perform conversion.
    #' @param show_formula If TRUE, the underlying Excel formulas are shown.
    #' @param convert If TRUE, a conversion to dates and numerics is attempted.
    #' @param skip_empty_cols If TRUE, empty columns are skipped.
    #' @param skip_empty_rows If TRUE, empty rows are skipped.
    #' @param skip_hidden_cols If TRUE, hidden columns are skipped.
    #' @param skip_hidden_rows If TRUE, hidden rows are skipped.
    #' @param start_row first row to begin looking for data.
    #' @param start_col first column to begin looking for data.
    #' @param rows A numeric vector specifying which rows in the Excel file to read. If NULL, all rows are read.
    #' @param cols A numeric vector specifying which columns in the Excel file to read. If NULL, all columns are read.
    #' @param named_region Character string with a named_region (defined name or table). If no sheet is selected, the first appearance will be selected.
    #' @param types A named numeric indicating, the type of the data. 0: character, 1: numeric, 2: date, 3: posixt, 4:logical. Names must match the returned data
    #' @param na.strings A character vector of strings which are to be interpreted as NA. Blank cells will be returned as NA.
    #' @param na.numbers A numeric vector of digits which are to be interpreted as NA. Blank cells will be returned as NA.
    #' @param fill_merged_cells If TRUE, the value in a merged cell is given to all cells within the merge.
    #' @param keep_attributes If TRUE additional attributes are returned. (These are used internally to define a cell type.)
    #' @param check_names If TRUE then the names of the variables in the data frame are checked to ensure that they are syntactically valid variable names.
    #' @param show_hyperlinks If `TRUE` instead of the displayed text, hyperlink targets are shown.
    #' @return a data frame
    to_df = function(
      sheet,
      start_row         = 1,
      start_col         = NULL,
      row_names         = FALSE,
      col_names         = TRUE,
      skip_empty_rows   = FALSE,
      skip_empty_cols   = FALSE,
      skip_hidden_rows  = FALSE,
      skip_hidden_cols  = FALSE,
      rows              = NULL,
      cols              = NULL,
      detect_dates      = TRUE,
      na.strings        = "#N/A",
      na.numbers        = NA,
      fill_merged_cells = FALSE,
      dims,
      show_formula      = FALSE,
      convert           = TRUE,
      types,
      named_region,
      keep_attributes   = FALSE,
      check_names       = FALSE,
      show_hyperlinks   = FALSE,
      ...
    ) {

      if (missing(sheet)) sheet <- substitute()
      if (missing(dims)) dims <- substitute()
      if (missing(named_region)) named_region <- substitute()

      standardize_case_names(...)

      wb_to_df(
        file              = self,
        sheet             = sheet,
        start_row         = start_row,
        start_col         = start_col,
        row_names         = row_names,
        col_names         = col_names,
        skip_empty_rows   = skip_empty_rows,
        skip_empty_cols   = skip_empty_cols,
        skip_hidden_rows  = skip_hidden_rows,
        skip_hidden_cols  = skip_hidden_cols,
        rows              = rows,
        cols              = cols,
        detect_dates      = detect_dates,
        na.strings        = na.strings,
        na.numbers        = na.numbers,
        fill_merged_cells = fill_merged_cells,
        dims              = dims,
        show_formula      = show_formula,
        convert           = convert,
        types             = types,
        named_region      = named_region,
        keep_attributes   = keep_attributes,
        check_names       = check_names,
        show_hyperlinks   = show_hyperlinks,
        ...               = ...
      )
    },

    ### load workbook ----
    #' @description load workbook
    #' @param file file
    #' @param data_only data_only
    #' @return The `wbWorkbook` object invisibly
    load = function(
      file,
      sheet,
      data_only  = FALSE,
      ...
    ) {
      # Is this required?
      if (missing(file))  file  <- substitute()
      if (missing(sheet)) sheet <- substitute()
      self <- wb_load(
        file       = file,
        sheet      = sheet,
        data_only  = data_only,
        ...        = ...
      )
      invisible(self)
    },

    # TODO wb_save can be shortened a lot by some formatting and by using a
    # function that creates all the temporary directories and subdirectries as a
    # named list

    #' @description
    #' Save the workbook
    #' @param file The path to save the workbook to
    #' @param overwrite If `FALSE`, will not overwrite when `path` exists
    #' @param path Deprecated argument previously used for file. Please use file in new code.
    #' @param flush Experimental, streams the worksheet file to disk
    #' @return The `wbWorkbook` object invisibly
    save = function(file = self$path, overwrite = TRUE, path = NULL, flush = FALSE) {

      if (!is.null(path)) {
        .Deprecated(old = "wb_save(path)", new = "wb_save(file)", package = "openxlsx2")
        file <- path
      }

      assert_class(file, "character")
      assert_class(overwrite, "logical")
      assert_class(flush, "logical")

      if (file.exists(file) & !overwrite) {
        stop("File already exists!")
      }

      valid_extensions <- c("xlsx", "xlsm") # "xlsb"
      file_extension   <- tolower(file_ext2(file))

      if (!file_extension %in% valid_extensions) {
        warning("The file extension '", file_extension,
        "' is invalid. Expected one of: ", paste0(valid_extensions, collapse = ", "),
        call. = FALSE)
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
      nTimelines      <- length(self$timelines)
      nComments       <- length(self$comments)
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
          con <- file(file.path(xlthemeDir, stringi::stri_join("theme", i, ".xml")), open = "wb")
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


        # TODO use seq_len() or seq_along()?
        for (i in seq_along(self$comments)) {

          if (length(self$comments[[i]]) && all(nchar(self$comments[[i]]))) {
            fn <- sprintf("comments%s.xml", i)

            write_comment_xml(
              comment_list = self$comments[[i]],
              file_name = file.path(tmpDir, "xl", fn)
            )
          }
        }

        private$writeDrawingVML(xldrawingsDir, xldrawingsRelsDir)
      }

      ## Threaded Comments xl/threadedComments/threadedComment.xml
      if (nThreadComments > 0) {
        xlThreadComments <- dir_create(tmpDir, "xl", "threadedComments")

        for (i in seq_along(self$threadComments)) {
          if (length(self$threadComments[[i]])) {
            write_file(
              head = "<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
              body = pxml(self$threadComments[[i]]),
              tail = "</ThreadedComments>",
              fl = file.path(xlThreadComments, sprintf("threadedComment%s.xml", i))
            )
          }
        }
      }

      ## xl/persons/person.xml
      if (nPersons) {
        personDir <- dir_create(tmpDir, "xl", "persons")
        write_file(
          body = self$persons,
          fl = file.path(personDir, "person.xml")
        )
      }

      ## xl/embeddings
      if (length(self$embeddings)) {
        embeddingsDir <- dir_create(tmpDir, "xl", "embeddings")
        for (fl in self$embeddings) {
          file.copy(fl, embeddingsDir, overwrite = TRUE)
        }
      }

      if (length(self$activeX)) {
        # we have to split activeX into activeX and activeX/_rels
        activeXDir     <- dir_create(tmpDir, "xl", "activeX")
        activeXRelsDir <- dir_create(tmpDir, "xl", "activeX", "_rels")
        for (fl in self$activeX) {
          if (file_ext2(fl) == "rels")
            file.copy(fl, activeXRelsDir, overwrite = TRUE)
          else
            file.copy(fl, activeXDir, overwrite = TRUE)
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

        for (i in seq_along(self$pivotTables)) {
          write_file(
            body = self$pivotTables[[i]],
            fl = file.path(pivotTablesDir, sprintf("pivotTable%s.xml", i))
          )
        }

        for (i in seq_along(self$pivotTables.xml.rels)) {
          write_file(
            body = self$pivotTables.xml.rels[[i]],
            fl = file.path(pivotTablesRelsDir, sprintf("pivotTable%s.xml.rels", i))
          )
        }

        for (i in seq_along(self$pivotRecords)) {
          write_file(
            body = self$pivotRecords[[i]],
            fl = file.path(pivotCacheDir, sprintf("pivotCacheRecords%s.xml", i))
          )
        }

        for (i in seq_along(self$pivotDefinitions)) {
          write_file(
            body = self$pivotDefinitions[[i]],
            fl = file.path(pivotCacheDir, sprintf("pivotCacheDefinition%s.xml", i))
          )
        }

        for (i in seq_along(self$pivotDefinitionsRels)) {
          write_file(
            body = self$pivotDefinitionsRels[[i]],
            fl = file.path(pivotCacheRelsDir, sprintf("pivotCacheDefinition%s.xml.rels", i))
          )
        }
      }

      ## slicers
      if (nSlicers) {
        slicersDir      <- dir_create(tmpDir, "xl", "slicers")
        slicerCachesDir <- dir_create(tmpDir, "xl", "slicerCaches")

        slicer_id <- which(self$slicers != "")
        for (i in slicer_id) {
          write_file(
            body = self$slicers[i],
            fl = file.path(slicersDir, sprintf("slicer%s.xml", i))
          )
        }

        caches_id <- which(self$slicerCaches != "")
        for (i in caches_id) {
          write_file(
            body = self$slicerCaches[[i]],
            fl = file.path(slicerCachesDir, sprintf("slicerCache%s.xml", i))
          )
        }
      }

      # timelines
      if (nTimelines) {
        timelinesDir      <- dir_create(tmpDir, "xl", "timelines")
        timelineCachesDir <- dir_create(tmpDir, "xl", "timelineCaches")

        timeline_id <- which(self$timelines != "")
        for (i in timeline_id) {
          write_file(
            body = self$timelines[i],
            fl = file.path(timelinesDir, sprintf("timeline%s.xml", i))
          )
        }

        caches_id <- which(self$timelineCaches != "")
        for (i in caches_id) {
          write_file(
            body = self$timelineCaches[[i]],
            fl = file.path(timelineCachesDir, sprintf("timelineCache%s.xml", i))
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

      # At the moment there is only a single known docMetadata file
      if (length(self$docMetadata)) {
        Ids <- c(Ids, paste0("rId", length(Ids) + 1L))
        Types <- c(
          Types,
          "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels"
        )
        Targets <- c(Targets, "docMetadata/LabelInfo.xml")
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

      ## write app.xml
      write_file(
        head = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
        body = pxml(self$app),
        tail = "</Properties>",
        fl = file.path(docPropsDir, "app.xml")
      )

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

        tab_ids <- table_ids(self)

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

      if (length(self$docMetadata)) {
        docMetadataDir     <- dir_create(tmpDir, "docMetadata")

        write_file(body = self$docMetadata, fl = file.path(docMetadataDir, "LabelInfo.xml"))

        ct <- append(ct, '<Override PartName="/docMetadata/LabelInfo.xml" ContentType="application/vnd.ms-office.classificationlabels+xml"/>')
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

      # featurePropertyBag
      if (length(self$featurePropertyBag)) {
        featurePropertyBagDir <- dir_create(tmpDir, "xl", "featurePropertyBag")

        write_file(
          body = self$featurePropertyBag,
          fl = file.path(
            featurePropertyBagDir,
            sprintf("featurePropertyBag.xml")
          )
        )
      }

      if (!is.null(self$richData)) {
        richDataDir <- dir_create(tmpDir, "xl", "richData")
        if (length(self$richData$richValueRel)) {
          write_file(
            body = self$richData$richValueRel,
            fl = file.path(
              richDataDir,
              "richValueRel.xml"
            )
          )
        }
        if (length(self$richData$rdrichvalue)) {
          write_file(
            body = self$richData$rdrichvalue,
            fl = file.path(
              richDataDir,
              "rdrichvalue.xml"
            )
          )
        }
        if (length(self$richData$rdrichvaluestr)) {
          write_file(
            body = self$richData$rdrichvaluestr,
            fl = file.path(
              richDataDir,
              "rdrichvaluestructure.xml"
            )
          )
        }
        if (length(self$richData$rdRichValueTypes)) {
          write_file(
            body = self$richData$rdRichValueTypes,
            fl = file.path(
              richDataDir,
              "rdRichValueTypes.xml"
            )
          )
        }

        if (length(self$richData$richValueRelrels)) {
          richDataRelDir <- dir_create(tmpDir, "xl", "richData", "_rels")
          write_file(
            body = self$richData$richValueRelrels,
            fl = file.path(
              richDataRelDir,
              "richValueRel.xml.rels"
            )
          )
        }
      }

      if (length(self$namedSheetViews)) {
        namedSheetViewsDir <- dir_create(tmpDir, "xl", "namedSheetViews")

        for (i in seq_along(self$namedSheetViews)) {
          write_file(body = self$namedSheetViews[i], fl = file.path(namedSheetViewsDir, sprintf("namedSheetView%i.xml", i)))
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
        xlworksheetsRelsDir,
        use_pugixml_export = isFALSE(flush)
      )

      ## write sharedStrings.xml
      if (length(self$sharedStrings)) {
        write_file(
          head = sprintf(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%s" uniqueCount="%s">',
            length(self$sharedStrings),
            attr(self$sharedStrings, "uniqueCount")
          ),
          body = stringi::stri_join(self$sharedStrings, collapse = "", sep = ""),
          tail = "</sst>",
          fl = file.path(xlDir, "sharedStrings.xml")
        )
      } else {
        ## Remove relationship to sharedStrings
        ct <- ct[!grepl("sharedStrings", ct)]
      }

      if (nComments > 0) {
        # FIXME why is this needed at all? We should not be required to modify Content_Types here ...
        # TODO This default extension is most likely wrong here and should be set when searching for and writing the vml entrys
        need_comments_xml <- which(self$comments != "")
        ct <- c(
          ct,
          '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
          sprintf('<Override PartName="/xl/comments%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>', need_comments_xml)
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
          stringi::stri_join(
            sprintf('<numFmts count="%s">', length(styleXML$numFmts)),
            pxml(styleXML$numFmts),
            "</numFmts>"
          )
      }
      styleXML$fonts <-
        stringi::stri_join(
          sprintf('<fonts count="%s">', length(styleXML$fonts)),
          pxml(styleXML$fonts),
          "</fonts>"
        )
      styleXML$fills <-
        stringi::stri_join(
          sprintf('<fills count="%s">', length(styleXML$fills)),
          pxml(styleXML$fills),
          "</fills>"
        )
      styleXML$borders <-
        stringi::stri_join(
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
        stringi::stri_join(
          sprintf('<cellXfs count="%s">', length(styleXML$cellXfs)),
          paste0(styleXML$cellXfs, collapse = ""),
          "</cellXfs>"
        )
      styleXML$cellStyles <-
        stringi::stri_join(
          sprintf('<cellStyles count="%s">', length(styleXML$cellStyles)),
          pxml(styleXML$cellStyles),
          "</cellStyles>"
        )

      styleXML$dxfs <-
        if (length(styleXML$dxfs)) {
          stringi::stri_join(
            sprintf('<dxfs count="%s">', length(styleXML$dxfs)),
            stringi::stri_join(unlist(styleXML$dxfs), sep = " ", collapse = ""),
            "</dxfs>"
          )
        } else {
          '<dxfs count="0"/>'
        }

      if (length(styleXML$tableStyles)) {
        styleXML$tableStyles <-
          xml_node_create(
            "tableStyles",
            xml_attributes = c(
              count = length(styleXML$tableStyles),
              defaultTableStyle = self$styles_mgr$defaultTableStyle,
              defaultPivotStyle = self$styles_mgr$defaultPivotStyle
            ),
            xml_children = styleXML$tableStyles
          )
      }

      # TODO
      # extLst

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
         head = "",
         body = '<styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
         tail = "",
         fl = file.path(xlDir, "styles.xml")
       )
      }

      if (length(self$calcChain)) {
        write_file(
          head = "",
          body = pxml(self$calcChain),
          tail = "",
          fl = file.path(xlDir, "calcChain.xml")
        )
      }

      # write metadata file. required if cm attribut is set.
      if (length(self$metadata)) {
        write_file(
          head = "",
          body = self$metadata,
          tail = "",
          fl = file.path(xlDir, "metadata.xml")
        )
      }

      ## write workbook.xml
      workbookXML <- self$workbook
      workbookXML$sheets <- stringi::stri_join("<sheets>", pxml(workbookXML$sheets), "</sheets>")

      if (length(workbookXML$definedNames)) {
        workbookXML$definedNames <- stringi::stri_join("<definedNames>", pxml(workbookXML$definedNames), "</definedNames>")
      }

      # openxml 2.8.1 expects the following order of xml nodes. While we create this per default, it is not
      # assured that the order of entries is still valid when we write the file. Functions can change the
      # workbook order, therefore we have to make sure that the expected order is written.
      # Otherwise spreadsheet software will complain.
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

      CT <- read_xml(paste0(tmpDir, "/[Content_Types].xml"))
      CT <- rbindlist(xml_attr(CT, "Types", "Override"))
      CT$tmpDirPartName <- paste0(tmpDir, CT$PartName)
      CT$fileExists <- file.exists(CT$tmpDirPartName)

      if (!all(CT$fileExists)) {
        missing_in_tmp <- CT$PartName[!CT$fileExists]
        warning(
          "[CT] file expected to be in output is missing: ",
          paste(missing_in_tmp, collapse = " ")
        )
      }

      WR <- read_xml(paste0(tmpDir, "/xl/_rels/workbook.xml.rels"))
      WR <- rbindlist(xml_attr(WR, "Relationships", "Relationship"))
      WR$tmpDirPartName <- paste0(tmpDir, "/xl/", WR$Target)
      WR$fileExists <- file.exists(WR$tmpDirPartName)

      if (!all(WR$fileExists)) {
        missing_in_tmp <- WR$Target[!WR$fileExists]
        warning(
          "[WR] file expected to be in output is missing: ",
          paste(missing_in_tmp, collapse = " ")
        )
      }

      folders <- c(
        # other tables to add?
        # pivotTables
        # embeddings
        "charts",
        "chartsheets",
        "drawings",
        "tables",
        "worksheets"
      )

      for (folder in folders) {

        path_ws_rels <- paste0(tmpDir, "/xl/", folder, "/_rels")

        ws_rels <- dir.exists(path_ws_rels)
        if (ws_rels) {
          # this somehow returned character(0)
          WR <- dir(path_ws_rels, full.names = TRUE)
          WR <- paste(vapply(WR, FUN = function(x) {
            paste(
              stringi::stri_read_lines(x, encoding = "UTF-8"),
              collapse = ""
            )
          }, FUN.VALUE = ""), collapse = "")
          if (WR != "") {
            WR <- rbindlist(xml_attr(WR, "Relationships", "Relationship"))

            if (NROW(WR)) { # in xlsb files it can be that WR has no rows
              WR$tmpDirPartName <- paste0(tmpDir, "/xl/", folder, "/", WR$Target)
              WR$fileExists <- file.exists(WR$tmpDirPartName)

              # exclude hyperlinks
              WR$type <- basename(WR$Type)
              WR <- WR[WR$type != "hyperlink", ]

              if (!all(WR$fileExists)) {
                missing_in_tmp <- WR$Target[!WR$fileExists]
                warning(
                  "[", folder, "] file expected to be in output is missing: ",
                  paste(missing_in_tmp, collapse = " ")
                )
              }
            }
          }
        }

      }


      # TODO make self$vbaProject be TRUE/FALSE
      tmpFile <- tempfile(tmpdir = tmpDir, fileext = if (isTRUE(self$vbaProject)) ".xlsm" else ".xlsx")

      # typo until release 1.8
      compr_level <- getOption("openxlsx2.compression_level") %||%
        getOption("openxlsx2.compresssionevel") %||%
        6L

      ## zip it
      zip::zip(
        zipfile = tmpFile,
        files = list.files(tmpDir, full.names = FALSE),
        recurse = TRUE,
        compression_level = compr_level,
        include_directories = FALSE,
        # change the working directory for this
        root = tmpDir,
        # change default to match historical zipr
        mode = "cherry-pick"
      )

      # Copy file; stop if failed
      if (!file.copy(from = tmpFile, to = file, overwrite = overwrite, copy.mode = FALSE)) {
        stop("Failed to save workbook")
      }

      # (re)assign file path (if successful)
      self$path <- file
      invisible(self)
    },

    #' @description open wbWorkbook in Excel.
    #' @details minor helper wrapping xl_open which does the entire same thing
    #' @param interactive If `FALSE` will throw a warning and not open the path.
    #'   This can be manually set to `TRUE`, otherwise when `NA` (default) uses
    #'   the value returned from [base::interactive()]
    #' @param flush flush
    #' @return The `wbWorkbook`, invisibly
    open = function(interactive = NA, flush = FALSE) {
      xl_open(self, interactive = interactive, flush = flush)
      invisible(self)
    },

    #' @description
    #' Build table
    #' @param colNames colNames
    #' @param ref ref
    #' @param showColNames showColNames
    #' @param tableStyle tableStyle
    #' @param tableName tableName
    #' @param withFilter withFilter
    #' @param totalsRowCount totalsRowCount
    #' @param totalLabel totalLabel
    #' @param showFirstColumn showFirstColumn
    #' @param showLastColumn showLastColumn
    #' @param showRowStripes showRowStripes
    #' @param showColumnStripes showColumnStripes
    #' @return The `wbWorksheet` object, invisibly
    buildTable = function(
      sheet             = current_sheet(),
      colNames,
      ref,
      showColNames,
      tableStyle,
      tableName,
      withFilter        = TRUE,
      totalsRowCount    = 0,
      totalLabel        = FALSE,
      showFirstColumn   = 0,
      showLastColumn    = 0,
      showRowStripes    = 1,
      showColumnStripes = 0
    ) {

      id <- as.character(last_table_id(self) + 1) # otherwise will start at 0 for table 1 length indicates the last known
      sheet <- private$get_sheet_index(sheet)
      # get the next highest rid
      rid <- 1
      if (!all(identical(self$worksheets_rels[[sheet]], character()))) {
        rid <- max(as.integer(sub("\\D+", "", rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))[["Id"]]))) + 1
      }

      if (is.null(self$tables)) {
        nms     <- NULL
        tSheets <- NULL
        tNames  <- NULL
        tActive <- NULL
      } else {
        nms     <- self$tables$tab_ref
        tSheets <- self$tables$tab_sheet
        tNames  <- self$tables$tab_name
        tActive <- self$tables$tab_act
      }


      ### autofilter
      autofilter <- if (withFilter) {
        autofilter_ref <- ref
        xml_node_create(xml_name = "autoFilter", xml_attributes = c(ref = autofilter_ref))
      }

      trf <- NULL
      has_total_row <- FALSE
      has_total_lbl <- FALSE
      if (!isFALSE(totalsRowCount)) {
        trf <- totalsRowCount
        has_total_row <- TRUE

        if (length(totalLabel) == length(colNames)) {
          lbl <- totalLabel
          has_total_lbl <- all(is.na(totalLabel))
        } else {
          lbl <- rep(NA_character_, length(colNames))
          has_total_lbl <- FALSE
        }

        rowcol   <- dims_to_rowcol(ref)
        ref_rows <- as.integer(rowcol[["row"]])
        ref      <- rowcol_to_dims(c(ref_rows, max(ref_rows) + 1L), rowcol[["col"]])
      }


      ### tableColumn
      tableColumn <- sapply(colNames, function(x) {
        id <- which(colNames %in% x)
        trf_id <- if (has_total_row) trf[[id]] else NULL
        lbl_id <- if (has_total_lbl && !is.na(lbl[[id]])) lbl[[id]] else NULL
        xml_node_create(
          "tableColumn",
          xml_attributes = c(
            id                = id,
            name              = x,
            totalsRowFunction = trf_id,
            totalsRowLabel    = lbl_id
          )
        )
      })

      tableColumns <- xml_node_create(
        xml_name       = "tableColumns",
        xml_children   = tableColumn,
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
        totalsRowCount = as_xml_attr(has_total_row),
        totalsRowShown = as_xml_attr(has_total_row)
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

    #' @description update a data_table
    #' @param tabname a tablename
    #' @return The `wbWorksheet` object, invisibly
    update_table = function(sheet = current_sheet(), dims = "A1", tabname) {

      sheet <- private$get_sheet_index(sheet)

      tabs <- self$get_tables(sheet = sheet)
      sel <- row.names(tabs[tabs$tab_name %in% tabname])

      wb_tabs <- self$tables[rownames(self$tables) %in% sel, ]

      xml <- wb_tabs$tab_xml
      tab_nams <- xml_node_name(xml, "table")
      known_xml <- c("autoFilter", "tableColumns", "tableStyleInfo")
      tab_unks <- tab_nams[!tab_nams %in% known_xml]
      if (length(tab_unks)) {
        msg <- paste(
          "Found unknown table xml nodes. These are lost using update_table: ",
          tab_unks
        )
        warning(msg)
      }

      tab_attr <- xml_attr(xml, "table")[[1]]
      tab_attr[["ref"]] <- dims

      tab_autofilter <- NULL
      if ("autofilter" %in% tab_nams) {
        tab_autofilter <- xml_node(xml, "table", "autoFilter")
        tab_autofilter <- xml_attr_mod(tab_autofilter, xml_attributes = c(ref = dims))
      }

      tab_tabColumns <- xml_node(xml, "table", "tableColumns")
      tab_cols <- names(self$to_df(sheet = sheet, dims = dims))

      tab_tabColumns <- fun_tab_cols(tab_cols)

      tab_tabStyleIn <- xml_node(xml, "table", "tableStyleInfo")

      xml <- xml_node_create(
        "table",
        xml_attributes = tab_attr,
        xml_children = c(
          tab_autofilter,
          tab_tabColumns,
          tab_tabStyleIn
        )
      )

      wb_tabs$tab_xml <- xml
      wb_tabs$tab_ref <- dims

      self$tables[rownames(self$tables) %in% sel, ] <- wb_tabs

      invisible(self)
    },

    ### copy cells ----

    #' @description
    #' copy cells around in a workbook
    #' @param data a wb_data object
    #' @param as_value should a copy of the value be written
    #' @param as_ref should references to the cell be written
    #' @param transpose should the data be written transposed
    #' @param ... additional arguments passed to add_data() if used with `as_value`
    #' @return The `wbWorksheet` object, invisibly
    copy_cells = function(
      sheet     = current_sheet(),
      dims      = "A1",
      data,
      as_value  = FALSE,
      as_ref    = FALSE,
      transpose = FALSE,
      ...
    ) {

      assert_class(data, "wb_data")
      from_sheet   <- attr(data, "sheet")
      from_dims_df <- attr(data, "dims")

      sheet <- private$get_sheet_index(sheet)

      to_ncol <- ncol(data) - 1
      to_nrow <- nrow(data) - 1

      start_col <- col2int(dims)
      start_row <- as.integer(gsub("\\D+", "", dims))

      to_cols <- seq.int(start_col, start_col + to_ncol)
      to_rows <- seq.int(start_row, start_row + to_nrow)

      if (transpose) {
        to_cols <- seq.int(start_col, start_col + to_nrow)
        to_rows <- seq.int(start_row, start_row + to_ncol)
        from_dims_df <- as.data.frame(t(from_dims_df), stringsAsFactors = FALSE)
      }

      to_dims       <- rowcol_to_dims(to_rows, to_cols)
      to_dims_df_i  <- dims_to_dataframe(to_dims, fill = FALSE)
      to_dims_df_f  <- dims_to_dataframe(to_dims, fill = TRUE)

      to_dims_f <- unname(unlist(to_dims_df_f))

      from_sheet <- private$get_sheet_index(from_sheet)
      from_dims  <- as.character(unlist(from_dims_df))
      cc <- self$worksheets[[from_sheet]]$sheet_data$cc

      # TODO improve this. It should use v or inlineStr from cc
      if (as_value) {
        data <- as.data.frame(unclass(data), stringsAsFactors = FALSE)

        if (transpose) {
          data <- t(data)
        }

        self$add_data(sheet = sheet, x = data, dims = to_dims_f[[1]], col_names = FALSE, ...)

        return(invisible(self))
      }

      # initialize dims we write to as empty cells
      private$do_cell_init(sheet, to_dims)

      to_cc <- cc[match(from_dims, cc$r), ]
      from_cells <- to_cc$r

      to_cc[c("r", "row_r", "c_r")] <- data.frame(
        r     = to_dims_f,
        row_r = gsub("\\D+", "", to_dims_f),
        c_r   = int2col(col2int(to_dims_f)),
        stringsAsFactors = FALSE
      )

      if (as_ref) {
        from_sheet_name <- self$get_sheet_names(escape = TRUE)[[from_sheet]]
        to_cc[names(to_cc) %in% c("c_t", "c_cm", "c_ph", "c_vm", "v", "f", "f_attr", "is")] <- ""
        to_cc[c("f")] <- paste0(shQuote(from_sheet_name, type = "sh"), "!", from_dims)
      }

      # uninitialized cells are NA_character_
      to_cc[is.na(to_cc)] <- ""

      cc <- self$worksheets[[sheet]]$sheet_data$cc
      cc[match(to_dims_f, cc$r), ] <- to_cc

      self$worksheets[[sheet]]$sheet_data$cc <- cc

      ### add hyperlinks ---
      if (length(self$worksheets[[from_sheet]]$relships$hyperlink)) {

        ws_hyls <- self$worksheets[[from_sheet]]$hyperlinks
        ws_rels <- self$worksheets_rels[[self$worksheets[[from_sheet]]$relships$hyperlink]]

        relships <- rbindlist(xml_attr(ws_rels, "Relationship"))
        relships <- relships[basename(relships$Type) == "hyperlink", ]

        # prepare hyperlinks data frame
        hlinks <- rbindlist(xml_attr(ws_hyls, "hyperlink"))

        # merge both
        hl_df <- merge(hlinks, relships, by.x = "r:id", by.y = "Id", all.x = TRUE, all.y = FALSE)

        hyperlink_in_wb <- hlinks$ref

        if (any(sel <- hyperlink_in_wb %in% from_dims)) {

          has_hl <- apply(from_dims_df, 2, function(x) x %in% hyperlink_in_wb)

          # are these always the same size?
          old <- from_dims_df[has_hl]
          new <- to_dims_df_f[has_hl]

          for (hls in match(hyperlink_in_wb, old)) {

            # prepare the updated link
            need_clone <- hyperlink_in_wb[hls]

            hl_df <- hlinks[hlinks$ref == need_clone, ]
            # this assumes that old and new are the same size
            hl_df$ref <- new[hls]
            hl <- df_to_xml("hyperlink", hl_df)

            # assign it
            self$worksheets[[sheet]]$hyperlinks <- append(
              self$worksheets[[sheet]]$hyperlinks,
              hl
            )
          }

        }

      }

      invisible(self)
    },

    ### base font ----

    #' @description Get the base font
    #' @return A list of of the font
    get_base_font = function() {
      baseFont <- self$styles_mgr$styles$fonts[[1]]

      sz     <- unlist(xml_attr(baseFont, "font", "sz"))
      color <- unlist(xml_attr(baseFont, "font", "color"))
      name   <- unlist(xml_attr(baseFont, "font", "name"))

      if (length(sz[[1]]) == 0) {
        sz <- list("val" = "11")
      } else {
        sz <- as.list(sz)
      }

      if (length(color[[1]]) == 0) {
        color <- list("rgb" = "#000000")
      } else {
        color <- as.list(color)
      }

      if (length(name[[1]]) == 0) {
        name <- list("val" = "Aptos Narrow")
      } else {
        name <- as.list(name)
      }

      list(
        size   = sz,
        color = color,
        name   = name
      )
    },

    #' @description Set the base font
    #' @param font_size fontSize
    #' @param font_color font_color
    #' @param font_name font_name
    #' @return The `wbWorkbook` object
    set_base_font = function(
      font_size  = 11,
      font_color = wb_color(theme = "1"),
      font_name  = "Aptos Narrow",
       ...
    ) {
      arguments <- c("font_size", "font_color", "font_name",
        "font_type", "font_panose")
      standardize(..., arguments = arguments)
      if (font_size < 0) stop("Invalid font_size")
      if (!is_wbColour(font_color)) font_color <- wb_color(font_color)

      fl <- system.file("extdata", "panose", "panose.csv", package = "openxlsx2")
      panose <- read.csv(fl, stringsAsFactors = FALSE)

      # if the default font name differes from the wanted name: update theme
      if (self$get_base_font()$name$val != font_name) {
        if (!exists("font_type")) font_type <- "Regular"

        sel <- panose$family == font_name & panose$type == font_type
        if (!any(sel) && !exists("font_panose")) {
          panose_hex <- NULL
        } else if (exists("font_panose")) {
          # the input provides a panose value
          panose_hex <- font_panose
        } else {
          panose_hex <- panose[sel, "panose"]
        }

        if (is.null(self$theme)) self$theme <- genBaseTheme()

        xml_font <- xml_node_create(
          "a:latin",
          xml_attributes = c(typeface = font_name, panose = panose_hex)
        )

        # TODO This alters both fonts. Should be able to alter indepdendently
        fS <- xml_node(self$theme, "a:theme", "a:themeElements", "a:fontScheme")
        maj_font <- xml_node(fS, "a:fontScheme", "a:majorFont", "a:latin")
        min_font <- xml_node(fS, "a:fontScheme", "a:minorFont", "a:latin")

        self$theme <- gsub(maj_font, xml_font, self$theme)
        self$theme <- gsub(min_font, xml_font, self$theme)
      }

      self$styles_mgr$styles$fonts[[1]] <- create_font(sz = font_size, color = font_color, name = font_name)
      invisible(self)
    },

    #' @description Get the base color
    #' @param xml xml
    #' @param plot plot
    get_base_colors = function(xml = FALSE, plot = TRUE) {

      if (is.null(self$theme)) self$theme <- genBaseTheme()

      current <- xml_node(self$theme, "a:theme", "a:themeElements", "a:clrScheme")
      name    <- xml_attr(current, "a:clrScheme")[[1]][["name"]]

      nodes  <- xml_node_name(current, "a:clrScheme")
      childs <- xml_node_name(current, "a:clrScheme", "*")

      rgbs <- vapply(
        seq_along(nodes),
        function(x) {
          nm <- nodes[x]
          cld <- childs[x]
          paste0("#", rbindlist(xml_attr(current, "a:clrScheme", nm, cld))[[1]])
        },
        NA_character_
      )
      names(rgbs) <- nodes

      if (interactive() && plot)
        barplot(
          rep(1, length(rgbs)),
          col = rgbs, names.arg = names(rgbs),
          main = name, yaxt = "n", las = 2
        )

      out <- list(rgbs)
      names(out) <- name

      if (xml) out <- current

      out
    },

    #' @description Get the base colour
    #' @param xml xml
    #' @param plot plot
    get_base_colours = function(xml = FALSE, plot = TRUE) {
      self$get_base_colors(xml = xml, plot = plot)
    },

    #' @description Set the base color
    #' @param theme theme
    #' @param ... ...
    #' @return The `wbWorkbook` object
    set_base_colors = function(theme = "Office", ...) {

      xml <- list(...)$xml

      if (is.null(xml)) {
        # read predefined themes
        clr_rds <- system.file("extdata", "colors.rds", package = "openxlsx2")
        colors <- readRDS(clr_rds)

        if (is.character(theme)) {
          sel <- match(theme, names(colors))
          err <- is.na(sel)
        } else {
          sel <- theme
          err <- sel > length(colors)
        }

        if (err) {
          stop("theme ", theme, " not found. doing nothing")
        }

        new <- colors[[sel]]
      } else {
        new <- xml
      }

      if (is.null(self$theme)) self$theme <- genBaseTheme()

      current    <- xml_node(self$theme, "a:theme", "a:themeElements", "a:clrScheme")
      self$theme <- stringi::stri_replace_all_fixed(self$theme, current, new)

      invisible(self)
    },

    #' @description Set the base colour
    #' @param theme theme
    #' @param ... ...
    #' @return The `wbWorkbook` object
    set_base_colours = function(theme = "Office", ...) {
      self$set_base_colors(theme = theme, ... = ...)
    },

    ### book views ----

    #' @description Get the book views
    #' @return A dataframe with the bookview properties
    get_bookview = function() {
      wbv <- self$workbook$bookViews
      if (is.null(wbv)) {
        wbv <- xml_node_create("workbookView")
      } else {
        wbv <- xml_node(wbv, "bookViews", "workbookView")
      }
      rbindlist(xml_attr(wbv, "workbookView"))
    },

    #' @description Get the book views
    #' @param view view
    #' @return The `wbWorkbook` object
    remove_bookview = function(view = NULL) {

      wbv <- self$workbook$bookViews

      if (is.null(wbv)) {
        return(invisible(self))
      } else {
        wbv <- xml_node(wbv, "bookViews", "workbookView")
      }

      if (!is.null(view)) {
        if (!is.integer(view)) view <- as.integer(view)
        # if there are three views, and 2 is removed, the indices are
        # now 1, 2 and not 1, 3. removing -1 keeps only the first view
        wbv <- wbv[-view]
      }

      self$workbook$bookViews <- xml_node_create(
        "bookViews",
        xml_children = wbv
      )

      invisible(self)
    },

    #' @param active_tab activeTab
    #' @param auto_filter_date_grouping autoFilterDateGrouping
    #' @param first_sheet firstSheet
    #' @param minimized minimized
    #' @param show_horizontal_scroll showHorizontalScroll
    #' @param show_sheet_tabs showSheetTabs
    #' @param show_vertical_scroll showVerticalScroll
    #' @param tab_ratio tabRatio
    #' @param visibility visibility
    #' @param window_height windowHeight
    #' @param window_width windowWidth
    #' @param x_window xWindow
    #' @param y_window yWindow
    #' @param view view
    #' @return The `wbWorkbook` object
    set_bookview = function(
      active_tab                = NULL,
      auto_filter_date_grouping = NULL,
      first_sheet               = NULL,
      minimized                 = NULL,
      show_horizontal_scroll    = NULL,
      show_sheet_tabs           = NULL,
      show_vertical_scroll      = NULL,
      tab_ratio                 = NULL,
      visibility                = NULL,
      window_height             = NULL,
      window_width              = NULL,
      x_window                  = NULL,
      y_window                  = NULL,
      view                      = 1L,
      ...
    ) {

      standardize_case_names(...)

      wbv <- self$workbook$bookViews

      if (is.null(wbv)) {
        wbv <- xml_node_create("workbookView")
      } else {
        wbv <- xml_node(wbv, "bookViews", "workbookView")
      }

      if (view > length(wbv)) {
        if (view == length(wbv) + 1L) {
          wbv <- c(wbv, xml_node_create("workbookView"))
        } else {
          msg <- paste0(
            "There is more than one workbook view missing.",
            " Available: ", length(wbv), ". Requested: ", view
          )
          stop(msg, call. = FALSE)
        }
      }

      wbv[view] <- xml_attr_mod(
        wbv[view],
        xml_attributes = c(
          activeTab              = as_xml_attr(active_tab),
          autoFilterDateGrouping = as_xml_attr(auto_filter_date_grouping),
          firstSheet             = as_xml_attr(first_sheet),
          minimized              = as_xml_attr(minimized),
          showHorizontalScroll   = as_xml_attr(show_horizontal_scroll),
          showSheetTabs          = as_xml_attr(show_sheet_tabs),
          showVerticalScroll     = as_xml_attr(show_vertical_scroll),
          tabRatio               = as_xml_attr(tab_ratio),
          visibility             = as_xml_attr(visibility),
          windowHeight           = as_xml_attr(window_height),
          windowWidth            = as_xml_attr(window_width),
          xWindow                = as_xml_attr(x_window),
          yWindow                = as_xml_attr(y_window)
        ),
        remove_empty_attr = FALSE
      )

      self$workbook$bookViews <- xml_node_create(
        "bookViews",
        xml_children = wbv
      )

      invisible(self)
    },

    ### sheet names ----

    #' @description Get sheet names
    #' @param escape Logical if the xml special characters are escaped
    #' @return A `named` `character` vector of sheet names in their order.  The
    #'   names represent the original value of the worksheet prior to any
    #'   character substitutions.
    get_sheet_names = function(escape = FALSE) {
      res <- private$original_sheet_names
      if (escape) res <- self$sheet_names
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
        return(invisible(self))
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
        return(invisible(self))
      }

      bad <- duplicated(tolower(new))
      if (any(bad)) {
        stop("Sheet names cannot have duplicates: ", toString(new[bad]))
      }

      # should be able to pull this out into a single private function
      for (i in seq_along(pos)) {
        sheet <- new_name[i]
        private$validate_new_sheet(sheet)
        new_name[i] <- sheet
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
          ind <- self$get_named_regions()
          # TODO why is the order switched?
          ind <- ind[order(as.integer(rownames(ind))), ]
          ind <- ind$sheets == old

          if (any(ind)) {
            nn <- sprintf("'%s'", new_name[i])
            nn <- stringi::stri_replace_all_fixed(self$workbook$definedNames[ind], old, nn)
            nn <- stringi::stri_replace_all(nn, regex = "'+", replacement = "'")
            self$workbook$definedNames[ind] <- nn
          }
        }
      }

      invisible(self)
    },


    ### row heights ----

    #' @description Sets a row height for a sheet
    #' @param rows rows
    #' @param heights heights
    #' @param hidden hidden
    #' @return The `wbWorkbook` object, invisibly
    set_row_heights = function(sheet = current_sheet(), rows, heights = NULL, hidden = FALSE) {
      sheet <- private$get_sheet_index(sheet)
      assert_class(heights, c("numeric", "integer"), or_null = TRUE, arg_nm = "heights")

      # TODO move to wbWorksheet method

      # create all A columns so that row_attr is available.
      # Someone thought that it would be a splendid idea, if
      # all row_attr needs to match cc. This is fine, though
      # it brings the downside that these cells have to be
      # initialized.
      dims <- rowcol_to_dims(rows, 1)
      private$do_cell_init(sheet, dims)

      row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr
      sel <- match(as.character(as.integer(rows)), row_attr$r)
      sel <- sel[!is.na(sel)]

      if (!is.null(heights)) {
        if (length(rows) > length(heights)) {
          heights <- rep_len(heights, length(rows))
        }

        if (length(heights) > length(rows)) {
          stop("Greater number of height values than rows.")
        }

        row_attr[sel, "ht"] <- as_xml_attr(heights)
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

    #' @description Removes a row height for a sheet
    #' @param rows rows
    #' @return The `wbWorkbook` object, invisibly
    remove_row_heights = function(sheet = current_sheet(), rows) {
      sheet <- private$get_sheet_index(sheet)

      row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

      if (is.null(row_attr)) {
        warning("There are no initialized rows on this sheet")
        return(invisible(self))
      }

      sel <- match(as.character(as.integer(rows)), row_attr$r)
      sel <- sel[!is.na(sel)]
      row_attr[sel, "ht"] <- ""
      row_attr[sel, "customHeight"] <- ""

      self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr

      invisible(self)
    },

    ## columns ----

    #' @description creates column object for worksheet
    #' @param n n
    #' @param beg beg
    #' @param end end
    createCols = function(sheet = current_sheet(), n, beg, end) {
       sheet <- private$get_sheet_index(sheet)
       self$worksheets[[sheet]]$cols_attr <- df_to_xml("col", empty_cols_attr(n, beg, end))
    },

    #' @description Group cols
    #' @param cols cols
    #' @param collapsed collapsed
    #' @param levels levels
    #' @return The `wbWorkbook` object, invisibly
    group_cols = function(sheet = current_sheet(), cols, collapsed = FALSE, levels = NULL) {
      sheet <- private$get_sheet_index(sheet)

      sPr <- self$worksheets[[sheet]]$sheetPr
      xml_sumRig <- unlist(xml_attr(sPr, "sheetPr", "outlinePr"))["summaryRight"]

      if (!is.null(xml_sumRig) && xml_sumRig == "0")
        right <- FALSE
      else
        right <- TRUE

      if (is.list(cols)) {
        cols <- lapply(cols, function(x) {
          if (is.list(x)) lapply(x, col2int)
          else col2int(x)
        })
        unis <- unique(unlist(cols))
        levels <- vector("character", length(unis))

        lvls <- names(cols)
        for (lvl in lvls) {
          grp_col_lvls <- cols[[lvl]]
          if (!is.list(grp_col_lvls)) grp_col_lvls <- list(grp_col_lvls)
          for (grp_col in grp_col_lvls) {
            collapse_in <- ifelse(right, length(grp_col), 1)
            sel <- unis %in% grp_col[-collapse_in]
            levels[sel] <- lvl
          }
        }
        cols <- unlist(cols)
      } else {
        cols <- col2int(cols)
        levels <- levels %||% rep("1", length(cols))
        collapse_in <- ifelse(right, length(levels), 1)
        levels[collapse_in] <- ""
      }

      if (length(collapsed) > length(cols)) {
        stop("Collapses argument is of greater length than number of cols.")
      }

      if (!is.logical(collapsed)) {
        stop("Collapses should be a logical value (TRUE/FALSE).")
      }

      if (any(cols < 1L)) {
        stop("Invalid rows entered (<= 0).")
      }

      # all collapsed = TRUE
      hidden <- all(collapsed)
      collapsed <- rep_len(as.character(as.integer(collapsed)), length(cols))

      # Remove duplicates
      ok <- !duplicated(cols)
      collapsed <- collapsed[ok]
      levels    <- levels[ok]
      cols      <- cols[ok]

      # create empty cols
      col_attr <- wb_create_columns(self, sheet, cols)

      # get the selection based on the col_attr frame.

      # the first n -1 cols get outlineLevel
      select <- col_attr$min %in% as.character(cols)
      collapse_in <- ifelse(right, length(cols), 1)
      select_n1 <- col_attr$min %in% as.character(cols[-collapse_in])
      if (length(select)) {
        col_attr$outlineLevel[select] <- as.character(levels)
        col_attr$collapsed[select] <- as_binary(collapsed)
        col_attr$hidden[select_n1] <- as_binary(hidden)
      }

      self$worksheets[[sheet]]$fold_cols(col_attr)


      # check if there are valid outlineLevel in col_attr and assign outlineLevelRow the max outlineLevel (thats in the documentation)
      if (any(col_attr$outlineLevel != "")) {
        self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(
          self$worksheets[[sheet]]$sheetFormatPr,
          xml_attributes = c(outlineLevelCol = as.character(max(as.integer(col_attr$outlineLevel), na.rm = TRUE))))
      }

      invisible(self)
    },

    #' @description ungroup cols
    #' @param cols columns
    #' @return The `wbWorkbook` object
    ungroup_cols = function(sheet = current_sheet(), cols) {
      sheet <- private$get_sheet_index(sheet)

      # check if any rows are selected
      if (any(cols < 1L)) {
        stop("Invalid cols entered (<= 0).")
      }

      # fetch the cols_attr data.frame
      col_attr <- self$worksheets[[sheet]]$unfold_cols()

      # get the selection based on the col_attr frame.
      select <- col_attr$min %in% as.character(col2int(cols))

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
    #' @description Set column widths
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
        return(invisible(self))
      }

      cols <- col2int(cols)

      if (length(widths) > length(cols)) {
        stop("More widths than columns supplied.")
      }

      if (length(hidden) > length(cols)) {
        stop("hidden argument is longer than cols.")
      }

      compatible_length <- length(cols) %% length(widths) == 0

      if (!compatible_length) {
        # needed because rep(c(1, 2 ), length.out = 3) is successful,
        # but not clear if this is what the user wanted.
        warning("`cols` and `widths` should have compatible lengths.\n",
             "`cols` has length ", length(cols), " while ",
             "`widths` has length ", length(widths), ".")
      }

      if (length(widths) < length(cols)) {
        widths <- rep_len(widths, length(cols))
      }
      compatible_length <- length(cols) %% length(hidden) == 0

      if (!compatible_length) {
        warning("`cols` and `hidden` should have compatible lengths.\n",
             "`cols` has length ", length(cols), " while ",
             "`hidden` has length ", length(hidden), ".")
      }

      if (length(hidden) < length(cols)) {
        hidden <- rep_len(hidden, length(cols))
      }

      # TODO add bestFit option?
      bestFit <- rep_len("1", length(cols))
      customWidth <- rep_len("1", length(cols))

      ## Remove duplicates
      ok <- !duplicated(cols)
      col_width <- widths[ok]
      hidden <- hidden[ok]
      cols <- cols[ok]

      base_font <- wb_get_base_font(self)

      if (any(widths == "auto")) {
        df <- wb_to_df(self, sheet = sheet, cols = cols, col_names = FALSE, keep_attributes = TRUE)
        # exclude merged cells from width calculation.
        # adapted from wb_to_df(fill_merged_cells = TRUE)
        mc <- self$worksheets[[sheet]]$mergeCells
        if (length(mc)) {
          mc <- unlist(xml_attr(mc, "mergeCell"))

          for (i in seq_along(mc)) {
            filler <- stringi::stri_split_fixed(mc[i], pattern = ":")[[1]][1]

            dms <- dims_to_dataframe(mc[i])

            if (any(row_sel <- rownames(df) %in% rownames(dms)) &&
                any(col_sel <- colnames(df) %in% colnames(dms))) {

              df[row_sel,  col_sel] <- NA
            }
          }
        }
        # TODO format(x) might not be the way it is formatted in the xlsx file.

        # With wb_to_df(col_names = FALSE) double values are not converted
        # from character to double. This was resulting in very wide columns.
        # To avoid this, we have to check character vectors for potentially
        # unconverted numerics and have to apply something like format(as.numeric(...))
        tt  <- attr(df, "tt")
        sel <- tt == "n" & !is.na(df)

        df[sel]   <- vapply(df[sel], function(x) format(as.numeric(x)), NA_character_)
        col_width <- vapply(df, function(x) max(nchar(format(x))), NA_real_)

        # add one extra spacing
        col_width <- col_width + 1L
      }


      # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column
      widths <- calc_col_width(base_font = base_font, col_width = col_width)

      # create empty cols
      col_df <- wb_create_columns(self, sheet, cols)

      select <- as.numeric(col_df$min) %in% cols
      col_df$width[select] <- as_xml_attr(widths)
      col_df$hidden[select] <- tolower(hidden)
      col_df$bestFit[select] <- bestFit
      col_df$customWidth[select] <- customWidth
      self$worksheets[[sheet]]$fold_cols(col_df)
      invisible(self)
    },

    ## rows ----

    # TODO groupRows() and groupCols() are very similiar.  Can problem turn
    # these into some wrappers for another method

    #' @description Group rows
    #' @param rows rows
    #' @param collapsed collapsed
    #' @param levels levels
    #' @return The `wbWorkbook` object, invisibly
    group_rows = function(sheet = current_sheet(), rows, collapsed = FALSE, levels = NULL) {
      sheet <- private$get_sheet_index(sheet)

      sPr <- self$worksheets[[sheet]]$sheetPr
      xml_sumBel <- unlist(xml_attr(sPr, "sheetPr", "outlinePr"))["summaryBelow"]

      if (!is.null(xml_sumBel) && xml_sumBel == "0") {
        below <- FALSE
      } else {
        below <- TRUE
      }

      if (is.list(rows)) {
        unis <- unique(unlist(rows))
        levels <- vector("character", length(unis))

        lvls <- names(rows)
        for (lvl in lvls) {
          grp_row_lvls <- rows[[lvl]]
          if (!is.list(grp_row_lvls)) grp_row_lvls <- list(grp_row_lvls)
          for (grp_row in grp_row_lvls) {
            collapse_in <- ifelse(below, length(grp_row), 1)
            sel <- unis %in% grp_row[-collapse_in]
            levels[sel] <- lvl
          }
        }
        rows <- unlist(rows)
      } else {
        levels <- levels %||% rep("1", length(rows))
        collapse_in <- ifelse(below, length(levels), 1)
        levels[collapse_in] <- ""
      }

      if (length(collapsed) > length(rows)) {
        stop("Collapses argument is of greater length than number of rows.")
      }

      if (!is.logical(collapsed)) {
        stop("Collapses should be a logical value (TRUE/FALSE).")
      }

      if (any(rows <= 0L)) {
        stop("Invalid rows entered (<= 0).")
      }

      # all collapsed = TRUE
      hidden <- all(collapsed)
      collapsed <- rep_len(as.character(as.integer(collapsed)), length(rows))

      # Remove duplicates
      ok <- !duplicated(rows)
      collapsed <- collapsed[ok]
      levels <- levels[ok]
      rows <- rows[ok]
      sheet <- private$get_sheet_index(sheet)

      # check if additional rows are required
      has_rows <- sort(as.integer(self$worksheets[[sheet]]$sheet_data$row_attr$r))
      missing_rows <- rows[!rows %in% has_rows]
      if (length(missing_rows)) private$do_cell_init(sheet, paste0("A", sort(missing_rows)))

      # fetch the row_attr data.frame
      row_attr <- self$worksheets[[sheet]]$sheet_data$row_attr

      # get the selection based on the row_attr frame.

      # the first n -1 rows get outlineLevel
      select <- row_attr$r %in% as.character(as.integer(rows))
      collapse_in <- ifelse(below, length(rows), 1)
      select_n1 <- row_attr$r %in% as.character(rows[-collapse_in])
      if (length(select)) {
        row_attr$outlineLevel[select] <- as.character(levels)
        row_attr$collapsed[select] <- as_binary(collapsed)
        row_attr$hidden[select_n1] <- as_binary(hidden)
      }

      self$worksheets[[sheet]]$sheet_data$row_attr <- row_attr

      # check if there are valid outlineLevel in row_attr and assign outlineLevelRow the max outlineLevel (thats in the documentation)
      if (any(row_attr$outlineLevel != "")) {
        self$worksheets[[sheet]]$sheetFormatPr <- xml_attr_mod(
          self$worksheets[[sheet]]$sheetFormatPr,
          xml_attributes = c(outlineLevelRow = as.character(max(as.integer(row_attr$outlineLevel), na.rm = TRUE))))
      }

      invisible(self)
    },

    #' @description ungroup rows
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
      select <- row_attr$r %in% as.character(as.integer(rows))
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

    #' @description Remove a worksheet
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

        # Removing these is probably a bad idea
        # NULL the sheets
        comment_id    <- self$worksheets[[sheet]]$relships$comments
        drawing_id    <- self$worksheets[[sheet]]$relships$drawing
        thrComment_id <- self$worksheets[[sheet]]$relships$threadComments
        vmlDrawing_id <- self$worksheets[[sheet]]$relships$vmlDrawing
        if (length(comment_id))    self$comments[[comment_id]]          <- ""
        if (length(drawing_id))    self$drawings[[drawing_id]]          <- ""
        if (length(drawing_id))    self$drawings_rels[[drawing_id]]     <- ""
        if (length(thrComment_id)) self$threadComments[[thrComment_id]] <- ""
        if (length(vmlDrawing_id)) self$vml[[vmlDrawing_id]]            <- ""
        if (length(vmlDrawing_id)) self$vml_rels[[vmlDrawing_id]]       <- ""

        #### Modify Content_Types
        ## remove drawing
        drawing_name <- xml_rels$target[xml_rels$type == "drawings"]
        if (!is.null(drawing_name) && !identical(drawing_name, character()))
          self$Content_Types <- grep(drawing_name, self$Content_Types, invert = TRUE, value = TRUE)

        # remove comment
        comment_name <- xml_rels$target[xml_rels$type == "comments"]
        if (!is.null(comment_name) && !identical(comment_name, character()))
          self$Content_Types <- grep(comment_name, self$Content_Types, invert = TRUE, value = TRUE)

      }

      nCharts <- max(which(self$is_chartsheet), 0)
      nWorks  <- max(which(!self$is_chartsheet), nSheets)

      self$is_chartsheet <- self$is_chartsheet[-sheet]

      ## remove the sheet
      ct <- read_Content_Types(self$Content_Types)
      self$Content_Types <- write_Content_Types(ct, rm_sheet = sheet)

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

        toRemove <- stringi::stri_join(
          sprintf("(pivotCacheDefinition%i\\.xml)", fileNo),
          sep = " ",
          collapse = "|"
        )

        toRemove <- stringi::stri_join(
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

      if (any(grepl("timelines", self$worksheets_rels[[sheet]]))) {
        # don't change to a grep(value = TRUE)
        self$workbook.xml.rels <- self$workbook.xml.rels[!grepl(sprintf("(timelineCache%s\\.xml)", sheet), self$workbook.xml.rels)]
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

      ## remove sheet
      sn <- apply_reg_match0(self$workbook$sheets, pat = '(?<= name=")[^"]+')
      self$workbook$sheets <- self$workbook$sheets[!sn %in% sheet_names]

      ## Reset rIds
      if (nSheets > 1) {
        for (i in (sheet + 1L):nSheets) {
          self$workbook$sheets <- gsub(
            stringi::stri_join("rId", i),
            stringi::stri_join("rId", i - 1L),
            self$workbook$sheets,
            fixed = TRUE
          )
          # these are zero indexed
          self$workbook$bookViews <- gsub(
            stringi::stri_join("activeTab=\"", i - 1L, "\""),
            stringi::stri_join("activeTab=\"", i - 2L, "\""),
            self$workbook$bookViews,
            fixed = TRUE
          )
        }
      } else {
        self$workbook$sheets <- NULL
      }

      wxr <- read_workbook.xml.rels(self$workbook.xml.rels)
      self$workbook.xml.rels <- write_workbook.xml.rels(wxr, rm_sheet = sheet)

      invisible(self)
    },

    #' @description Adds data validation
    #' @param type type
    #' @param operator operator
    #' @param value value
    #' @param allow_blank allowBlank
    #' @param show_input_msg showInputMsg
    #' @param show_error_msg showErrorMsg
    #' @param error_style The icon shown and the options how to deal with such inputs. Default "stop" (cancel), else "information" (prompt popup) or "warning" (prompt accept or change input)
    #' @param error_title The error title
    #' @param error The error text
    #' @param prompt_title The prompt title
    #' @param prompt The prompt text
    #' @return The `wbWorkbook` object
    add_data_validation = function(
      sheet          = current_sheet(),
      dims           = "A1",
      type,
      operator,
      value,
      allow_blank    = TRUE,
      show_input_msg = TRUE,
      show_error_msg = TRUE,
      error_style    = NULL,
      error_title    = NULL,
      error          = NULL,
      prompt_title   = NULL,
      prompt         = NULL,
      ...
    ) {

      sheet <- private$get_sheet_index(sheet)

      cols <- list(...)[["cols"]]
      rows <- list(...)[["rows"]]

      if (!is.null(rows) && !is.null(cols)) {
        .Deprecated(old = "cols/rows", new = "dims", package = "openxlsx2")
        dims <- rowcol_to_dims(rows, cols)
      }

      standardize(...)

      assert_class(allow_blank, "logical")
      assert_class(show_input_msg, "logical")
      assert_class(show_error_msg, "logical")

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

      # prepare for worksheet
      origin <- get_date_origin(self, origin = TRUE)

      if (type == "list") {
        operator <- NULL
      }

      self$worksheets[[sheet]]$.__enclos_env__$private$data_validation(
        type         = type,
        operator     = operator,
        value        = value,
        allowBlank   = as_xml_attr(allow_blank),
        showInputMsg = as_xml_attr(show_input_msg),
        showErrorMsg = as_xml_attr(show_error_msg),
        errorStyle   = error_style,
        errorTitle   = error_title,
        error        = error,
        promptTitle  = prompt_title,
        prompt       = prompt,
        origin       = origin,
        sqref        = dims
      )

      invisible(self)
    },

    #' @description
    #' Set cell merging for a sheet
    #' @param solve logical if intersecting cells should be solved
    #' @param direction direction in which to split the cell merging. Allows "row" or "col".
    #' @return The `wbWorkbook` object, invisibly
    merge_cells = function(sheet = current_sheet(), dims = NULL, solve = FALSE, direction = NULL, ...) {

      cols <- list(...)[["cols"]]
      rows <- list(...)[["rows"]]

      if (!is.null(rows) && !is.null(cols)) {
        .Deprecated(old = "cols/rows", new = "dims", package = "openxlsx2")
        dims <- rowcol_to_dims(rows, cols)
      }

      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$merge_cells(dims = dims, solve = solve, direction = direction)
      invisible(self)
    },

    #' @description
    #' Removes cell merging for a sheet
    #' @return The `wbWorkbook` object, invisibly
    unmerge_cells = function(sheet = current_sheet(), dims = NULL, ...) {

      cols <- list(...)[["cols"]]
      rows <- list(...)[["rows"]]

      if (!is.null(rows) && !is.null(cols)) {
        .Deprecated(old = "cols/rows", new = "dims", package = "openxlsx2")
        dims <- rowcol_to_dims(rows, cols)
      }

      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$unmerge_cells(
        dims = dims
      )
      invisible(self)
    },

    #' @description
    #' Set freeze panes for a sheet
    #' @param first_active_row first_active_row
    #' @param first_active_col first_active_col
    #' @param first_row first_row
    #' @param first_col first_col
    #' @return The `wbWorkbook` object, invisibly
    freeze_pane = function(
      sheet            = current_sheet(),
      first_active_row = NULL,
      first_active_col = NULL,
      first_row        = FALSE,
      first_col        = FALSE,
      ...
    ) {

      # TODO rename to setFreezePanes?
      standardize_case_names(...)

      # fine to do the validation before the actual check to prevent other errors
      sheet <- private$get_sheet_index(sheet)

      if (is.null(first_active_row) & is.null(first_active_col) & !first_row & !first_col) {
        return(invisible(self))
      }

      # TODO simplify asserts
      if (!is.logical(first_row)) stop("first_row must be TRUE/FALSE")
      if (!is.logical(first_col)) stop("first_col must be TRUE/FALSE")

      # make overwrides for arguments
      if (first_row & !first_col) {
        first_active_col <- NULL
        first_active_row <- NULL
        first_col <- FALSE
      } else if (first_col & !first_row) {
        first_active_row <- NULL
        first_active_col <- NULL
        first_row <- FALSE
      } else if (first_row & first_col) {
        first_active_row <- 2L
        first_active_col <- 2L
        first_row <- FALSE
        first_col <- FALSE
      } else {
        ## else both firstRow and firstCol are FALSE
        first_active_row <- first_active_row %||% 1L
        first_active_col <- first_active_col %||% 1L

        # Convert to numeric if column letter given
        # TODO is col2int() safe for non characters?
        first_active_row <- col2int(first_active_row)
        first_active_col <- col2int(first_active_col)
      }

      paneNode <-
        if (first_row) {
          '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        } else if (first_col) {
          '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>'
        } else {
          if (first_active_row == 1 & first_active_col == 1) {
            ## nothing to do
            # return(NULL)
            return(invisible(self))
          }

          if (first_active_row > 1 & first_active_col == 1) {
            attrs <- sprintf('ySplit="%s"', first_active_row - 1L)
            activePane <- "bottomLeft"
          }

          if (first_active_row == 1 & first_active_col > 1) {
            attrs <- sprintf('xSplit="%s"', first_active_col - 1L)
            activePane <- "topRight"
          }

          if (first_active_row > 1 & first_active_col > 1) {
            attrs <- sprintf('ySplit="%s" xSplit="%s"',
              first_active_row - 1L,
              first_active_col - 1L
            )
            activePane <- "bottomRight"
          }

          sprintf(
            '<pane %s topLeftCell="%s" activePane="%s" state="frozen"/><selection pane="%s"/>',
            stringi::stri_join(attrs, collapse = " ", sep = " "),
            get_cell_refs(data.frame(first_active_row, first_active_col, stringsAsFactors = FALSE)),
            activePane,
            activePane
          )
        }

      self$worksheets[[sheet]]$freezePane <- paneNode
      invisible(self)
    },

    ## comment ----

    #' @description Add comment
    #' @param dims row and column as spreadsheet dimension, e.g. "A1"
    #' @param comment a comment to apply to the worksheet
    #' @return The `wbWorkbook` object
    add_comment = function(
        sheet   = current_sheet(),
        dims    = "A1",
        comment,
        ...
    ) {

      col   <- list(...)[["col"]]
      row   <- list(...)[["row"]]
      color <- list(...)[["color"]] %||% list(...)[["colour"]]
      file  <- list(...)[["file"]]

      if (!is.null(row) && !is.null(col)) {
        .Deprecated(old = "col/row", new = "dims", package = "openxlsx2")
        dims <- rowcol_to_dim(row, col)
      }

      if (is.character(comment)) {
        comment <- wb_comment(text = comment, author = getOption("openxlsx2.creator"))
      }

      if (!is.null(color) && !is_wbColour(color))
        stop("color needs to be a wb_color()")

      do_write_comment(
        wb      = self,
        sheet   = sheet,
        comment = comment,
        dims    = dims,
        color   = color,
        file    = file
      ) # has no use: xy

      invisible(self)
    },

    #' @description Get comments
    #' @param sheet sheet
    #' @param dims dims
    #' @return A data frame containing comments
    get_comment = function(
      sheet = current_sheet(),
      dims  = NULL
    ) {

      sheet_id <- private$get_sheet_index(sheet)
      cmmt <- self$worksheets[[sheet_id]]$relships$comments

      if (!is.null(dims) && any(grepl(":", dims)))
        dims <- unname(unlist(dims_to_dataframe(dims, fill = TRUE)))

      cmts <- list()
      if (length(cmmt) && length(self$comments) <= cmmt) {
        cmts <- do.call(rbind, lapply(self$comments[[cmmt]], function(x) {
          data.frame(
            ref = x$ref,
            author = x$author,
            comment = paste0(x$comment, collapse = " "),
            stringsAsFactors = FALSE
          )
        }))

        if (!is.null(dims)) cmts <- cmts[cmts$ref %in% dims, ]
        # print(cmts)
        cmts <- cmts[c("ref", "author", "comment")]
        if (NROW(cmts)) {
          cmts$comment <- as_fmt_txt(cmts$comment)
          cmts$cmmt_id <- cmmt
        } else {
          return(NULL)
        }
      }

      cmts
    },

    #' @description Remove comment
    #' @param dims row and column as spreadsheet dimension, e.g. "A1"
    #' @return The `wbWorkbook` object
    remove_comment = function(
      sheet      = current_sheet(),
      dims       = "A1",
      ...
    ) {

      col <- list(...)[["col"]]
      row <- list(...)[["row"]]
      gridExpand <- list(...)[["gridExpand"]]

      if ((!is.null(row) && !is.null(col))) {
        .Deprecated(old = "col/row/gridExpand", new = "dims", package = "openxlsx2")
        dims <- wb_dims(row, col)
      }

      # TODO: remove with deprication
      if (is.null(gridExpand)) {
        # default until deprecating
        gridExpand <- TRUE
      }

      do_remove_comment(
        wb         = self,
        sheet      = sheet,
        col        = col,
        row        = row,
        dims       = dims,
        gridExpand = gridExpand
      )

      invisible(self)
    },


    #' @description add threaded comment to worksheet
    #' @param comment the comment to add
    #' @param person_id the person Id this should be added for
    #' @param reply logical if the comment is a reply
    #' @param resolve logical if the comment should be marked as resolved
    #' @export
    add_thread = function(
      sheet      = current_sheet(),
      dims       = "A1",
      comment    = NULL,
      person_id,
      reply      = FALSE,
      resolve    = FALSE
    ) {

      if (missing(person_id)) {
        person_id <- getOption("openxlsx2.thread_id")
        if (is.null(person_id)) stop("no person id found")
      }

      sheet <- private$get_sheet_index(sheet)
      wb_cmt <- wb_get_comment(self, sheet, dims)

      if (length(cmt <- wb_cmt$comment)) {
        # TODO not sure yet what to do
      } else {
        cmt <- wb_comment(text = comment, author = "")
        self$add_comment(sheet = sheet, dims = dims, comment = cmt)
      }
      wb_cmt <- wb_get_comment(self, sheet, dims)

      if (!length(self$worksheets[[sheet]]$relships$threadedComment)) {

        thread_id <- length(self$threadComments) + 1L

        # TODO the sheet id is correct ... ?
        self$worksheets[[sheet]]$relships$threadedComment <- thread_id

        self$append(
          "Content_Types",
          sprintf("<Override PartName=\"/xl/threadedComments/threadedComment%s.xml\" ContentType=\"application/vnd.ms-excel.threadedcomments+xml\"/>", thread_id)
        )

        self$worksheets_rels[[sheet]] <- append(
          self$worksheets_rels[[sheet]],
          sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment\" Target=\"../threadedComments/threadedComment%s.xml\"/>", length(self$worksheets_rels[[sheet]]) + 1L, thread_id)
        )

        self$threadComments[[thread_id]] <- character()
      }

      thread_id <- self$worksheets[[sheet]]$relships$threadedComment

      parentId <- NULL
      tcs <- rbindlist(xml_attr(self$threadComments[[thread_id]], "threadedComment"))
      sel <- which(tcs$ref == dims)

      if (reply && nrow(tcs)) {
        if (length(sel))  {
          parentId <- tcs[sel[1], ]$id
        } else {
          warning("cannot reply, will create a new thread")
        }
      }

      # update or remove any previous thread from the dims
      if (length(sel)) {
        if (resolve) {
          self$threadComments[[thread_id]][sel[1]] <- xml_attr_mod(
            self$threadComments[[thread_id]][sel[1]],
            xml_attributes = c(done = as_xml_attr(resolve))
          )
        } else if (!reply) {
          self$threadComments[[thread_id]] <- self$threadComments[[thread_id]][-(sel)]
        }
      }

      if (!is.null(comment)) {

        # For replies we can update the comment, but the id remains the parentId
        cmt_id <- st_guid()

        done <- as_xml_attr(resolve)
        if (reply) done <- NULL

        # Internal option to alleviate timing problems in CI and CRAN
        ts <- getOption("openxlsx2.datetimeCreated", default = Sys.time())

        tc <- xml_node_create(
          "threadedComment",
          xml_attributes = c(
            ref      = dims,
            dT       = format(as_POSIXct_utc(ts), "%Y-%m-%dT%H:%M:%SZ"),
            personId = person_id,
            id       = cmt_id,
            parentId = parentId,
            done     = done
          ),
          xml_children = xml_node_create("text", xml_children = comment)
        )

        self$threadComments[[thread_id]] <- append(
          self$threadComments[[thread_id]],
          tc
        )

        if (reply) cmt_id <- parentId

        wb_cmt <- wb_get_comment(self, sheet, dims)
        sId <- wb_cmt$cmmt_id
        cId <- as.integer(rownames(wb_cmt))

        tc <- cbind(
          rbindlist(xml_attr(self$threadComments[[thread_id]], "threadedComment")),
          text = xml_value(self$threadComments[[thread_id]], "threadedComment", "text"),
          stringsAsFactors = FALSE
        )

        # probably correclty ordered, but we could order these by date?
        tc <- tc[which(tc$ref == dims), ]

        tc <- paste0(
          "<t>[Threaded comment]\n\nYour spreadsheet software allows you to read this threaded comment; ",
          "however, any edits to it will get removed if the file is opened in a newer version of a certain spreadsheet software.\n\n",
          paste("Comment:", paste0(tc$text, collapse = "\nReplie:")),
          "</t>"
        )

        self$comments[[sId]][[cId]] <- list(
          ref = dims,
          author = sprintf("tc=%s", cmt_id),
          comment = tc,
          style = FALSE,
          clientData = NULL
        )

      }

      invisible(self)
    },

    #' @description Get threads
    #' @param sheet sheet
    #' @param dims dims
    #' @return A data frame containing threads
    get_thread = function(sheet = current_sheet(), dims = NULL) {

      sheet <- private$get_sheet_index(sheet)
      thrd <- self$worksheets[[sheet]]$relships$threadedComment

      tc <- cbind(
        rbindlist(xml_attr(self$threadComments[[thrd]], "threadedComment")),
        text = xml_value(self$threadComments[[thrd]], "threadedComment", "text"),
        stringsAsFactors = FALSE
      )

      if (!is.null(dims) && any(grepl(":", dims)))
        dims <- unname(unlist(dims_to_dataframe(dims, fill = TRUE)))

      if (!is.null(dims)) {
        tc <- tc[tc$ref %in% dims, ]
      }

      persons <- self$get_person()

      tc <- merge(tc, persons, by.x = "personId", by.y = "id",
                  all.x = TRUE, all.y = FALSE)

      tc$dT <- as.POSIXct(tc$dT, format = "%Y-%m-%dT%H:%M:%SZ")

      tc[c("dT", "ref", "displayName", "text", "done")]
    },

    ## conditional formatting ----

    #' @description Add conditional formatting
    #' @param rule rule
    #' @param style style
    #' @param type type
    #' @param params Additional parameters
    #' @return The `wbWorkbook` object
    add_conditional_formatting = function(
        sheet  = current_sheet(),
        dims   = NULL,
        rule   = NULL,
        style  = NULL,
        # TODO add vector of possible values
        type   = c(
          "expression", "colorScale",
          "dataBar", "iconSet",
          "duplicatedValues", "uniqueValues",
          "containsErrors", "notContainsErrors",
          "containsBlanks", "notContainsBlanks",
          "containsText", "notContainsText",
          "beginsWith", "endsWith",
          "between", "topN", "bottomN"
        ),
        params = list(
          showValue = TRUE,
          gradient  = TRUE,
          border    = TRUE,
          percent   = FALSE,
          rank      = 5L
        ),
        ...
    ) {

      cols <- list(...)[["cols"]]
      rows <- list(...)[["rows"]]

      if (!is.null(rows) && !is.null(cols)) {
        .Deprecated(old = "cols/rows", new = "dims", package = "openxlsx2")
        dims <- rowcol_to_dims(rows, cols)
      }

      # as_integer returns a range, but we want to know all columns
      ddims <- dims_to_rowcol(dims, as_integer = FALSE)
      rows <- sort(as.integer(ddims[["row"]]))
      cols <- sort(col2int(ddims[["col"]]))

      if (!is.null(style)) assert_class(style, "character")
      assert_class(type, "character")
      assert_class(params, "list")

      type <- match.arg(type)

      ## check valid rule
      dxfId <- NULL
      if (!is.null(style)) dxfId <- self$styles_mgr$get_dxf_id(style)
      params <- validate_cf_params(params)
      values <- NULL

      sel <- c("expression", "duplicatedValues", "containsText", "notContainsText", "beginsWith",
               "endsWith", "between", "topN", "bottomN", "uniqueValues", "iconSet",
               "containsErrors", "notContainsErrors", "containsBlanks", "notContainsBlanks")
      if (is.null(style) && type %in% sel) {
        smp <- random_string()
        style <- create_dxfs_style(font_color = wb_color(hex = "FF9C0006"), bg_fill = wb_color(hex = "FFFFC7CE"))
        self$styles_mgr$add(style, smp)
        dxfId <- self$styles_mgr$get_dxf_id(smp)
      }


      cols <- tapply(cols, cumsum(c(1, diff(cols) != 1)), function(g) {
        range(g)
      })

      rows <- tapply(rows, cumsum(c(1, diff(rows) != 1)), function(g) {
        range(g)
      })

      orig_rule <- rule

      for (row in rows) {
        for (col in cols) {


          switch(
            type,

            expression = {
              # TODO should we bother to do any conversions or require the text
              # entered to be exactly as an Excel expression would be written?
              msg <- "When type == 'expression', "

              if (!is.character(orig_rule) || length(orig_rule) != 1L) {
                stop(msg, "rule must be a single length character vector")
              }

              rule <- orig_rule

              rule <- gsub("!=", "<>", rule)
              rule <- gsub("==", "=", rule)
              rule <- replace_legal_chars(rule) # replaces <>

              if (!grepl("[A-Z]", substr(rule, 1, 2))) {
                ## formula looks like "operatorX" , attach top left cell to rule
                rule <- paste0(
                  get_cell_refs(data.frame(row[1], col[1], stringsAsFactors = FALSE)),
                  rule
                )
              } ## else, there is a letter in the formula and apply as is

            },

            colorScale = {
              # - style is a vector of colors with length 2 or 3
              # - rule specifies the quantiles (numeric vector of length 2 or 3), if NULL min and max are used
              msg <- "When type == 'colorScale', "

              if (!is.character(style)) {
                stop(msg, "style must be a vector of colors of length 2 or 3.")
              }

              if (!length(style) %in% 2:3) {
                stop(msg, "style must be a vector of length 2 or 3.")
              }

              if (!is.null(rule)) {
                if (length(rule) != length(style)) {
                  stop(msg, "rule and style must have equal lengths.")
                }
              }

              style <- validate_color(style)

              if (isFALSE(style)) {
                stop(msg, "style must be valid colors")
              }

              values <- rule
              rule <- style
            },

            dataBar = {
              # - style is a vector of colors of length 2 or 3
              # - rule specifies the quantiles (numeric vector of length 2 or 3), if NULL min and max are used
              msg <- "When type == 'dataBar', "
              style <- style %||% "#638EC6"

              # TODO use inherits() not class()
              if (!inherits(style, "character")) {
                stop(msg, "style must be a vector of colors of length 1 or 2.")
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
              style <- validate_color(style)

              if (isFALSE(style)) {
                stop(msg, "style must be valid colors")
              }

              values <- rule
              rule <- style
            },

            iconSet = {
              # - rule is the iconSet values
              msg <- "When type == 'iconSet', "
              values <- rule
            },

            duplicatedValues = {
              # type == "duplicatedValues"
              # - style is a Style object
              # - rule is ignored

              rule <- style
            },

            uniqueValues = {
              # type == "uniqueValues"
              # - style is a Style object
              # - rule is ignored

              rule <- style
            },

            containsBlanks = {
              # - style is Style object
              # - rule is cell to check for errors
              msg <- "When type == 'containsBlanks', "

              rule <- style
            },

            notContainsBlanks = {
              # - style is Style object
              # - rule is cell to check for errors
              msg <- "When type == 'notContainsBlanks', "

              rule <- style
            },

            containsErrors = {
              # - style is Style object
              # - rule is cell to check for errors
              msg <- "When type == 'containsErrors', "

              rule <- style
            },

            notContainsErrors = {
              # - style is Style object
              # - rule is cell to check for errors
              msg <- "When type == 'notContainsErrors', "

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
            startRow = row[1],
            endRow   = row[2],
            startCol = col[1],
            endCol   = col[2],
            dxfId    = dxfId,
            formula  = rule,
            type     = type,
            values   = values,
            params   = params
          )
        }
      }

      invisible(self)
    },

    #' @description Remove conditional formatting
    #' @param sheet sheet
    #' @param dims dims
    #' @param first first
    #' @param last last
    #' @return The `wbWorkbook` object
    remove_conditional_formatting = function(
        sheet  = current_sheet(),
        dims   = NULL,
        first  = FALSE,
        last   = FALSE
    ) {

      sheet <- private$get_sheet_index(sheet)

      if (is.null(dims) && isFALSE(first) && isFALSE(last)) {
        self$worksheets[[sheet]]$conditionalFormatting <- character()
      } else {

        cf <- self$worksheets[[sheet]]$conditionalFormatting

        if (!is.null(dims)) {
          if (any(sel <- names(cf) %in% dims)) {
            self$worksheets[[sheet]]$conditionalFormatting <- cf[!sel]
          }
        } else if (first) {
            self$worksheets[[sheet]]$conditionalFormatting <- cf[-1]
        } else if (last) {
            self$worksheets[[sheet]]$conditionalFormatting <- cf[-length(cf)]
        }
      }

      invisible(self)
    },

    ## plots and images ----

    #' @description
    #' Insert an image into a sheet
    #' @param file file
    #' @param width width
    #' @param height height
    #' @param row_offset,col_offset offsets
    #' @param units units
    #' @param dpi dpi
    #' @param address address
    #' @return The `wbWorkbook` object, invisibly
    add_image = function(
      sheet      = current_sheet(),
      dims       = "A1",
      file,
      width      = 6,
      height     = 3,
      row_offset = 0,
      col_offset = 0,
      units      = "in",
      dpi        = 300,
      address    = NULL,
      ...
    ) {

      arguments <- c(ls(), "start_row", "start_col")
      standardize_case_names(..., arguments = arguments)

      params <- list(...)
      if (!is.null(params$start_row)) start_row <- params$start_row
      if (!is.null(params$start_col)) start_col <- params$start_col

      if (exists("start_row") || exists("start_col")) {
        if (!exists("start_row")) start_row <- 1
        if (!exists("start_col")) start_col <- 1
        .Deprecated(old = "start_col/start_row", new = "dims", package = "openxlsx2")
        start_col <- col2int(start_col)
        start_row <- as.integer(start_row)
        dims <- rowcol_to_dim(start_row, start_col)
      }

      if (!file.exists(file)) {
        stop("File ", file, " does not exist.")
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

      private$add_media(file)
      file <- names(self$media)[length(self$media)]

      if (length(self$worksheets[[sheet]]$relships$drawing)) {
        sheet_drawing <- self$worksheets[[sheet]]$relships$drawing
        imageNo <- length(xml_node_name(self$drawings[[sheet_drawing]], "xdr:wsDr")) + 1L
      } else {
        sheet_drawing <- length(self$drawings) + 1L
        # self$append("drawings", NA_character_)
        # self$append("drawings_rels", "")
        imageNo <- 1L
      }

      # TODO might want to clean this a bit more
      if (is.null(address)) address_id <- ""

      if (length(self$drawings_rels) >= sheet_drawing && !all(self$drawings_rels[[sheet_drawing]] == "")) {
        next_id <- get_next_id(self$drawings_rels[[sheet_drawing]])
        if (!is.null(address)) address_id <- get_next_id(self$drawings_rels[[sheet_drawing]], 2L)
      } else {
        next_id <- "rId1"
        if (!is.null(address)) address_id <- "rId2"
      }

      pos <- '<xdr:pos x="0" y="0" />'

      drawingsXML <- stringi::stri_join(
        "<xdr:absoluteAnchor>",
        pos,
        sprintf('<xdr:ext cx="%s" cy="%s"/>', width, height),
        genBasePic(imageNo, next_id, address_id),
        "<xdr:clientData/>",
        "</xdr:absoluteAnchor>"
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

      self$add_drawing(
        sheet      = sheet,
        dims       = dims,
        xml        = drawing,
        col_offset = col_offset,
        row_offset = row_offset
      )

      # add image to drawings_rels
      old_drawings_rels <- unlist(self$drawings_rels[[sheet_drawing]])
      if (all(is.na(old_drawings_rels)) || all(old_drawings_rels == ""))
        old_drawings_rels <- NULL

      ## drawings rels (Reference from drawings.xml to image file in media folder)
      self$drawings_rels[[sheet_drawing]] <- c(
        old_drawings_rels,
        sprintf(
          '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/%s"/>',
          next_id,
          file
        )
      )

      if (!is.null(address)) {
        relship <- sprintf(
          '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>',
          address_id,
          address
        )
        self$drawings_rels[[sheet_drawing]] <- append(
          self$drawings_rels[[sheet_drawing]],
          relship
        )
      }

      invisible(self)
    },

    #' @description Add plot. A wrapper for add_image()
    #' @param width width
    #' @param height height
    #' @param row_offset,col_offset offsets
    #' @param file_type fileType
    #' @param units units
    #' @param dpi dpi
    #' @return The `wbWorkbook` object
    add_plot = function(
      sheet      = current_sheet(),
      dims       = "A1",
      width      = 6,
      height     = 4,
      row_offset = 0,
      col_offset = 0,
      file_type  = "png",
      units      = "in",
      dpi        = 300,
      ...
    ) {

      arguments <- c(ls(), "start_row", "start_col")
      standardize_case_names(..., arguments = arguments)

      params <- list(...)
      if (!is.null(params$start_row)) start_row <- params$start_row
      if (!is.null(params$start_col)) start_col <- params$start_col

      if (exists("start_row") || exists("start_col")) {
        if (!exists("start_row")) start_row <- 1
        if (!exists("start_col")) start_col <- 1
        .Deprecated(old = "start_row/start_col", new = "dims", package = "openxlsx2")
        dims <- rowcol_to_dim(start_row, start_col)
      }

      if (is.null(dev.list()[[1]])) {
        warning("No plot to insert.")
        return(invisible(self))
      }

      fileType <- tolower(file_type)
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

      fileName <- tempfile(pattern = "figureImage", fileext = paste0(".", file_type))

      # Workaround for wrapper test. Otherwise tempfile names differ

      if (identical(Sys.getenv("TESTTHAT"), "true")) fileName <- getOption("openxlsx2.temp_png")

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
      stopifnot(file.exists(fileName))

      self$add_image(
        sheet      = sheet,
        dims       = dims,
        file       = fileName,
        width      = width,
        height     = height,
        row_offset = row_offset,
        col_offset = col_offset,
        units      = units,
        dpi        = dpi
      )
    },

    #' @description Add xml drawing
    #' @param xml xml
    #' @param col_offset,row_offset offsets for column and row
    #' @return The `wbWorkbook` object
    add_drawing = function(
      sheet      = current_sheet(),
      dims       = "A1",
      xml,
      col_offset = 0,
      row_offset = 0,
      ...
    ) {

      edit_as <- NULL
      standardize_case_names(...)
      if (!is.null(list(...)$edit_as)) edit_as <- list(...)$edit_as

      sheet <- private$get_sheet_index(sheet)

      is_chartsheet <- self$is_chartsheet[sheet]

      # usually sheet_drawing is sheet. If we have userShapes, sheet_drawing
      # can skip ahead. see test: unemployment-nrw202208.xlsx
      if (length(self$worksheets[[sheet]]$relships$drawing)) {
        sheet_drawing <- self$worksheets[[sheet]]$relships$drawing

        # chartsheets can not have multiple drawings
        if (is_chartsheet) {
          self$drawings[[sheet_drawing]]      <- ""
          self$drawings_rels[[sheet_drawing]] <- ""
        }
      } else {
        sheet_drawing <- length(self$drawings) + 1L
        self$append("drawings", "")
        self$append("drawings_rels", "")
      }

      # prepare mschart drawing
      if (inherits(xml, "chart_id")) {
        xml <- drawings(self$drawings_rels[[sheet_drawing]], xml)
      }

      xml <- read_xml(xml, pointer = FALSE)

      if (!(xml_node_name(xml) == "xdr:wsDr")) {
        stop("xml needs to be a drawing.")
      }

      altc  <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "mc:AlternateContent")
      ext   <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:ext")
      pic   <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:pic")
      grpSp <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:grpSp")
      grFrm <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:graphicFrame")
      sp    <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:sp")
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
          if (length(col_offset) != 2) col_offset <- rep(col_offset, 2)
          if (length(row_offset) != 2) row_offset <- rep(row_offset, 2)

          anchor <- paste0(
            "<xdr:from>",
            "<xdr:col>%s</xdr:col><xdr:colOff>%s</xdr:colOff>",
            "<xdr:row>%s</xdr:row><xdr:rowOff>%s</xdr:rowOff>",
            "</xdr:from>",
            "<xdr:to>",
            "<xdr:col>%s</xdr:col><xdr:colOff>%s</xdr:colOff>",
            "<xdr:row>%s</xdr:row><xdr:rowOff>%s</xdr:rowOff>",
            "</xdr:to>"
          )
          anchor <- sprintf(
            anchor,
            cols[1] - 1L, col_offset[1],
            rows[1] - 1L, row_offset[1],
            cols[2], col_offset[2],
            rows[2], row_offset[2]
          )

        } else {

          xdr_typ <- "xdr:oneCellAnchor"

          cols <- col2int(dims)
          rows <- as.numeric(gsub("\\D+", "", dims))

          anchor <- paste0(
            "<xdr:from>",
            "<xdr:col>%s</xdr:col><xdr:colOff>%s</xdr:colOff>",
            "<xdr:row>%s</xdr:row><xdr:rowOff>%s</xdr:rowOff>",
            "</xdr:from>"
          )
          anchor <- sprintf(
            anchor,
            cols[1] - 1L, col_offset[1],
            rows[1] - 1L, row_offset[1]
          )

        }

        xdr_typ_xml <- xml_node_create(
          xdr_typ,
          xml_children = c(
            anchor,
            altc,
            ext,
            pic,
            grpSp,
            grFrm,
            sp,
            clDt
          ),
          xml_attributes = c(
            editAs = as_xml_attr(edit_as)
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

      # check if sheet already contains drawing. if yes, try to integrate
      # our drawing into this else we only use our drawing.
      drawings <- self$drawings[[sheet_drawing]]
      if (is.null(drawings) || is.na(drawings) || drawings == "") {
        drawings <- xml
      } else {
        drawing_type <- xml_node_name(xml, "xdr:wsDr")
        xml_drawing <- xml_node(xml, "xdr:wsDr", drawing_type)
        drawings <- xml_add_child(drawings, xml_drawing)
      }
      self$drawings[[sheet_drawing]] <- drawings

      self$worksheets[[sheet]]$relships$drawing <- sheet_drawing

      # get the correct next free relship id
      if (length(self$worksheets_rels[[sheet]]) == 0) {
        next_relship <- 1
        has_no_drawing <- TRUE
      } else {
        relship <- rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))
        relship$typ <- basename(relship$Type)
        next_relship <- max(as.integer(gsub("\\D+", "", relship$Id))) + 1L
        has_no_drawing <- !any(relship$typ == "drawing")
      }

      # if a drawing exisits, we already added ourself to it. Otherwise we
      # create a new drawing.
      if (has_no_drawing) {
        self$worksheets_rels[[sheet]] <- append(
          self$worksheets_rels[[sheet]],
          sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing%s.xml\"/>", next_relship, sheet_drawing)
        )
        self$worksheets[[sheet]]$drawing <- sprintf("<drawing r:id=\"rId%s\"/>", next_relship)
      }

      invisible(self)
    },

    #' @description Add xml chart
    #' @param xml xml
    #' @param col_offset,row_offset positioning parameters
    #' @return The `wbWorkbook` object
    add_chart_xml = function(
      sheet      = current_sheet(),
      dims       = NULL,
      xml,
      col_offset = 0,
      row_offset = 0,
      ...
    ) {

      arguments <- c(ls(), "start_row", "start_col")
      standardize_case_names(..., arguments = arguments)

      params <- list(...)
      if (!is.null(params$start_row)) start_row <- params$start_row
      if (!is.null(params$start_col)) start_col <- params$start_col

      if (exists("start_row") || exists("start_col")) {
        if (!exists("start_row")) start_row <- 1
        if (!exists("start_col")) start_col <- 1
        .Deprecated(old = "start_col/start_row", new = "dims", package = "openxlsx2")
        dims <- rowcol_to_dim(start_row, start_col)
      }

      sheet <- private$get_sheet_index(sheet)
      if (length(self$worksheets[[sheet]]$relships$drawing)) {
        # if one is found: we have to select this drawing
        sheet_drawing <- self$worksheets[[sheet]]$relships$drawing
      } else {
        # if none is found. we need to add a new drawing
        sheet_drawing <- length(self$drawings) + 1L
      }

      next_chart <- NROW(self$charts) + 1

      chart <- data.frame(
        chart   = xml,
        colors  = colors1_xml,
        style   = styleplot_xml,
        rels    = chart1_rels_xml(next_chart),
        chartEx = "",
        relsEx  = "",
        stringsAsFactors = FALSE
      )

      self$charts <- rbind(self$charts, chart)

      class(next_chart) <- c("integer", "chart_id")

      # create drawing. add it to self$drawings, the worksheet and rels
      self$add_drawing(
        sheet      = sheet,
        dims       = dims,
        xml        = next_chart,
        col_offset = col_offset,
        row_offset = row_offset
      )

      sheet_drawing <- self$worksheets[[sheet]]$relships$drawing

      self$drawings_rels[[sheet_drawing]] <- drawings_rels(
        self$drawings_rels[[sheet_drawing]],
        next_chart
      )

      invisible(self)
    },

    #' @description Add mschart chart to the workbook
    #' @param dims the dimensions where the sheet will appear
    #' @param graph mschart graph
    #' @param col_offset,row_offset offsets for column and row
    #' @return The `wbWorkbook` object
    add_mschart = function(
      sheet      = current_sheet(),
      dims       = NULL,
      graph,
      col_offset = 0,
      row_offset = 0,
      ...
    ) {

      standardize_case_names(...)

      requireNamespace("mschart")
      assert_class(graph, "ms_chart")

      sheetname <- private$get_sheet_name(sheet)

      # format.ms_chart is exported in mschart >= 0.4
      ids <- random_string(n = 2, length = 8, pattern = "[1-9]")
      out_xml <- read_xml(
        format(
          graph,
          sheetname = sheetname,
          id_x = ids[1],
          id_y = ids[2]
        ),
        pointer = FALSE
      )

      # write the chart data to the workbook
      if (inherits(graph$data_series, "wb_data")) {
        self$
          add_chart_xml(
            sheet      = sheet,
            dims       = dims,
            xml        = out_xml,
            col_offset = col_offset,
            row_offset = row_offset
          )
      } else {
        self$
          add_data(sheet = sheet, x = graph$data_series)$
          add_chart_xml(
            sheet      = sheet,
            dims       = dims,
            xml        = out_xml,
            col_offset = col_offset,
            row_offset = row_offset
          )
      }
    },

    #' @description Add form control to workbook
    #' @param type type
    #' @param text text
    #' @param link link
    #' @param range range
    #' @param checked checked
    #' @return The `wbWorkbook` object, invisibly
    add_form_control = function(
      sheet   = current_sheet(),
      dims    = "A1",
      type    = c("Checkbox", "Radio", "Drop"),
      text    = NULL,
      link    = NULL,
      range   = NULL,
      checked = FALSE
    ) {
      sheet <- private$get_sheet_index(sheet)

      if (!is.null(dims)) {
        xy <- dims_to_rowcol(dims)
        left <- col2int(xy[["col"]][1]) - 1L
        top  <- as.integer(xy[["row"]][1]) - 1L

        # for A1:B2
        if (length(xy[[1]]) > 1) {
          right  <- max(col2int(xy[["col"]]))
        } else {
          right  <- left + 1L
        }

        if (length(xy[[2]]) > 1) {
          bottom <- max(as.integer(xy[["row"]]))
        } else {
          bottom <- top + 1L
        }
      }

      text <- text %||% ""

      type <- match.arg(type)


      clientData <- genClientDataFC(left, top, right, bottom, link, range, type, checked)

      if (type == "Checkbox") {
        vml <- read_xml(
          sprintf(
            '<o:shapelayout v:ext="edit">
            <o:idmap v:ext="edit" data="1" />
            </o:shapelayout>
            <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">
            <v:stroke joinstyle="miter" />
            <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect" />
            <o:lock v:ext="edit" shapetype="t" />
            </v:shapetype>
            <v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute;  margin-left:57pt;margin-top:40pt;width:120pt;height:30pt;z-index:1;  mso-wrap-style:tight" filled="f" fillcolor="white [65]" stroked="f" strokecolor="black [64]" o:insetmode="auto">
            <v:path shadowok="t" strokeok="t" fillok="t" />
            <o:lock v:ext="edit" rotation="t" />
            <v:textbox style="mso-direction-alt:auto" o:singleclick="f">
            <div style="text-align:left">
            <font face="Lucida Grande" size="260" color="auto">%s</font>
            </div>
            </v:textbox>
            %s
            </v:shape>',
            text,
            clientData
          ), pointer = FALSE
        )
      } else if (type == "Radio") {
        vml <- read_xml(
          sprintf(
            '<v:shape id="_x0000_s1027" type="#_x0000_t201" style="position:absolute;  margin-left:69pt;margin-top:155pt;width:120pt;height:30pt;z-index:3;  mso-wrap-style:tight" filled="f" fillcolor="white [65]" stroked="f" strokecolor="black [64]" o:insetmode="auto">
            <v:path shadowok="t" strokeok="t" fillok="t" />
            <o:lock v:ext="edit" rotation="t" />
            <v:textbox style="mso-direction-alt:auto" o:singleclick="f">
            <div style="text-align:left">
            <font face="Lucida Grande" size="260" color="auto">%s</font>
            </div>
            </v:textbox>
            %s
            </v:shape>',
            text,
            clientData
          ), pointer = FALSE
        )
      } else if (type == "Drop") {
        vml <- read_xml(
          sprintf(
            '<v:shape id="_x0000_s1029" type="#_x0000_t201" style="position:absolute;  margin-left:336pt;margin-top:54pt;width:180pt;height:60pt;z-index:5" stroked="f" strokecolor="black [64]" o:insetmode="auto">
            <o:lock v:ext="edit" rotation="t" text="t" />
            %s
            </v:shape>',
            clientData
          ), pointer = FALSE
        )
      }


      # self$add_drawing(xml = drawing, dims = dims)
      vml_id <- self$worksheets[[sheet]]$relships$vmlDrawing

      if (is.null(unlist(self$vml[vml_id]))) {
        vml <- xml_node_create(
          "xml",
          xml_attributes = c(
            `xmlns:v` = "urn:schemas-microsoft-com:vml",
            `xmlns:o` = "urn:schemas-microsoft-com:office:office",
            `xmlns:x` = "urn:schemas-microsoft-com:office:excel"
          ),
          xml_children = c(
            vml
          )
        )
        self$append("vml", list(vml))
        self$worksheets[[sheet]]$relships$vmlDrawing <- length(self$vml)
      } else {
        self$vml[[vml_id]] <- xml_add_child(
          xml_node = self$vml[[vml_id]],
          xml_child = vml
        )
      }

      # wb$drawings

      drawing <- formCntrlDrawing(type, length(self$ctrlProps))

      self$add_drawing(sheet = sheet, xml = drawing, dims = dims)

      if (type == "Checkbox") {
        frmCntrl <- "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"CheckBox\" checked=\"Checked\" lockText=\"1\" noThreeD=\"1\"/>"
      } else if (type == "Radio") {
        frmCntrl <- "<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"Radio\" checked=\"Checked\" lockText=\"0\" noThreeD=\"1\"/>"
      } else if (type == "Drop") {
        frmCntrl <- '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Drop" dropStyle="combo" dx="15" noThreeD="1" sel="0" val="0"/>'
      }

      self$append(
        "ctrlProps",
        frmCntrl
      )

      ctrlProp <- length(self$ctrlProps)

      self$append(
        "Content_Types",
        sprintf(
          '<Override PartName="/xl/ctrlProps/ctrlProp%s.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>',
          ctrlProp
        )
      )

      # usually sheet_drawing is sheet. If we have userShapes, sheet_drawing
      # can skip ahead. see test: unemployment-nrw202208.xlsx
      found <- private$get_drawingsref()
      if (sheet %in% found$sheet) {
        sheet_drawing <- found$id[found$sheet == sheet]
      } else {
        sel <- which.min(abs(found$sheet - sheet))
        sheet_drawing <- max(sheet, found$id[found$sheet == sel] + 1)
      }

      # get the correct next free relship id
      if (length(self$worksheets_rels[[sheet]]) == 0) {
        next_relship <- 1
        has_no_drawing <- TRUE
        has_no_vmlDrawing <- TRUE
      } else {
        relship <- rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))
        relship$typ <- basename(relship$Type)
        next_relship <- as.integer(gsub("\\D+", "", relship$Id)) + 1L
        has_no_drawing <- !any(relship$typ == "drawing")
        has_no_vmlDrawing <- !any(relship$typ == "vmlDrawing")
      }

      if (has_no_vmlDrawing) {
        if (!any(grepl("vmlDrawing", self$Content_Types))) {
          self$append(
            "Content_Types",
            "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>"
          )
        }

        self$worksheets_rels[[sheet]] <- c(
          self$worksheets_rels[[sheet]],
          sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing%s.vml\"/>", next_relship, length(self$vml))
        )

        self$worksheets[[sheet]]$legacyDrawing <- sprintf("<legacyDrawing r:id=\"rId%s\"/>", next_relship)
      }

      invisible(self)
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
        sheetTxt <- sprintf("Sheets: %s", paste(exSheets, collapse = ", "))

        showText <- c(showText, sheetTxt, "\n")
      } else {
        showText <-
          c(showText, "\nWorksheets:\n", "No worksheets attached\n")
      }

      if (nSheets > 0) {
        showText <-
          c(showText, sprintf(
            "Write order: %s",
            stringi::stri_join(self$sheetOrder, sep = " ", collapse = ", ")
          ))
      }

      cat(unlist(showText))
      invisible(self)
    },

    #' @description
    #' Protect a workbook
    #' @param protect protect
    #' @param lock_structure lock_structure
    #' @param lock_windows lock_windows
    #' @param password password
    #' @param type type
    #' @param file_sharing file_sharing
    #' @param username username
    #' @param read_only_recommended read_only_recommended
    #' @return The `wbWorkbook` object, invisibly
    protect = function(
      protect               = TRUE,
      password              = NULL,
      lock_structure        = FALSE,
      lock_windows          = FALSE,
      type                  = 1,
      file_sharing          = FALSE,
      username              = unname(Sys.info()["user"]),
      read_only_recommended = FALSE,
      ...
    ) {

      standardize_case_names(...)

      if (!protect) {
        self$workbook$workbookProtection <- NULL
        return(invisible(self))
      }

      # match.arg() doesn't handle numbers too well
      type <- as_xml_attr(type)
      password <- if (is.null(password)) "" else hashPassword(password)

      # TODO: Shall we parse the existing protection settings and preserve all
      # unchanged attributes?

      if (file_sharing) {
        self$workbook$fileSharing <- xml_node_create(
          "fileSharing",
          xml_attributes = c(
            userName = username,
            readOnlyRecommended = if (read_only_recommended | type == "2") "1",
            reservationPassword = password
          )
        )
      }

      self$workbook$workbookProtection <- xml_node_create(
        "workbookProtection",
        xml_attributes = c(
          hashPassword = password,
          lockStructure = toString(as.numeric(lock_structure)),
          lockWindows = toString(as.numeric(lock_windows))
        )
      )

      self$app$DocSecurity <- xml_node_create("DocSecurity", xml_children = type)
      invisible(self)
    },

    #' @description protect worksheet
    #' @param protect protect
    #' @param password password
    #' @param properties A character vector of properties to lock.  Can be one
    #'   or more of the following: `"selectLockedCells"`,
    #'   `"selectUnlockedCells"`, `"formatCells"`, `"formatColumns"`,
    #'   `"formatRows"`, `"insertColumns"`, `"insertRows"`,
    #'   `"insertHyperlinks"`, `"deleteColumns"`, `"deleteRows"`, `"sort"`,
    #'   `"autoFilter"`, `"pivotTables"`, `"objects"`, `"scenarios"`
    #' @return The `wbWorkbook` object
    protect_worksheet = function(
        sheet = current_sheet(),
        protect    = TRUE,
        password   = NULL,
        properties = NULL
    ) {

      sheet <- private$get_sheet_index(sheet)

      if (!protect) {
        # initializes as character()
        self$worksheets[[sheet]]$sheetProtection <- character()
        return(invisible(self))
      }

      all_props <- worksheet_lock_properties()

      if (!is.null(properties)) {
        # ensure only valid properties are listed
        if (is.null(names(properties))) {
          properties <- match.arg(properties, all_props, several.ok = TRUE)
          properties <- as_xml_attr(all_props %in% properties)
          names(properties) <- all_props
          properties <- properties[properties != "0"]
        } else {
          keep <- match.arg(names(properties), all_props, several.ok = TRUE)
          properties <- properties[keep]
          nms <- names(properties)
          properties <- as_xml_attr(properties)
          names(properties) <- nms
        }
      }

      if (!is.null(password))
        properties <- c(properties, password = hashPassword(password))

      self$worksheets[[sheet]]$sheetProtection <- xml_node_create(
        "sheetProtection",
        xml_attributes = c(
          sheet = "1",
          properties
        )
      )

      invisible(self)
    },


    ### creators --------------------------------------------------------------

    #' @description Get properties of a workbook
    get_properties = function() {
      nams <- xml_node_name(self$core, "cp:coreProperties")
      properties <- vapply(nams, function(x) {
        xml_value(self$core, "cp:coreProperties", x, escapes = TRUE)
      },
      FUN.VALUE = NA_character_)

      name_replace <- c(
        title = "dc:title",
        subject = "dc:subject",
        creator = "dc:creator",
        keywords = "cp:keywords",
        comments = "dc:description",
        modifier = "cp:lastModifiedBy",
        datetime_created = "dcterms:created",
        datetime_modified = "dcterms:modified",
        category = "cp:category"
      )
      # use names
      names(properties) <- names(name_replace)[match(names(properties), name_replace)]


      if (!is.null(self$app$Company)) {
        properties <- c(properties, "company" = xml_value(self$app$Company, level1 = "Company"))
      }
      if (!is.null(self$app$Manager)) {
        properties <- c(properties, "manager" = xml_value(self$app$Manager, level1 = "Manager"))
      }
      properties
    },

    #' @description Set a property of a workbook
    #' @param title,subject,category,datetime_created,datetime_modified,modifier,keywords,comments,manager,company,custom A workbook property to set
    set_properties = function(
      creator           = NULL,
      title             = NULL,
      subject           = NULL,
      category          = NULL,
      datetime_created  = NULL,
      datetime_modified = NULL,
      modifier          = NULL,
      keywords          = NULL,
      comments          = NULL,
      manager           = NULL,
      company           = NULL,
      custom            = NULL
    ) {

      # Internal option to alleviate timing problems in CI and CRAN
      datetime_created <-
        getOption("openxlsx2.datetimeCreated", datetime_created)


      core_dctitle <- "dc:title"
      core_subject <- "dc:subject"
      core_creator <- "dc:creator"
      core_keyword <- "cp:keywords"
      core_describ <- "dc:description"
      core_lastmod <- "cp:lastModifiedBy"
      core_created <- "dcterms:created"
      core_modifid <- "dcterms:modified"
      core_categor <- "cp:category"

      # get an xml output or create one
      if (!is.null(self$core)) {
        nams <- xml_node_name(self$core, "cp:coreProperties")
        xml_properties <- vapply(nams, function(x) {
          xml_node(self$core, "cp:coreProperties", x, escapes = TRUE)
        }, FUN.VALUE = NA_character_)
      } else {
        xml_properties <- c(
          core_dctitle = "",
          core_subject = "",
          core_creator = "",
          core_keyword = "",
          core_describ = "",
          core_lastmod = "",
          core_created = "",
          core_modifid = "",
          core_categor = ""
        )
      }

      if (!is.null(title)) {
        xml_properties[core_dctitle] <- xml_node_create(core_dctitle, xml_children = title)
      }

      if (!is.null(subject)) {
        xml_properties[core_subject] <- xml_node_create(core_subject, xml_children = subject)
      }

      # update values where needed
      if (!is.null(creator)) {
        if (length(creator) > 1) creator <- paste0(creator, collapse = ";")
        xml_properties[core_creator] <- xml_node_create(core_creator, xml_children = creator)
        modifier <- creator
      }

      if (!is.null(keywords)) {
        xml_properties[core_keyword] <- xml_node_create(core_keyword, xml_children = keywords)
      }

      if (!is.null(comments)) {
        xml_properties[core_describ] <- xml_node_create(core_describ, xml_children = comments)
      }

      if (!is.null(manager)) {
        self$app$Manager <- xml_node_create("Manager", xml_children = manager)
      }

      if (!is.null(company)) {
        self$app$Company <- xml_node_create("Company", xml_children = company)
      }

      if (!is.null(datetime_created)) {
        xml_properties[core_created] <- xml_node_create(core_created,
          xml_attributes = c(
            `xsi:type` = "dcterms:W3CDTF"
          ),
          xml_children = format(as_POSIXct_utc(datetime_created), "%Y-%m-%dT%H:%M:%SZ")
        )
      }

      if (!is.null(datetime_modified)) {
        xml_properties[core_modifid] <- xml_node_create(core_modifid,
          xml_attributes = c(
            `xsi:type` = "dcterms:W3CDTF"
          ),
          xml_children = format(as_POSIXct_utc(datetime_modified), "%Y-%m-%dT%H:%M:%SZ")
        )
      }

      if (!is.null(modifier)) {
        xml_properties[core_lastmod] <- xml_node_create(core_lastmod, xml_children = modifier)
      }

      if (!is.null(category)) {
        xml_properties[core_categor] <- xml_node_create(core_categor, xml_children = category)
      }

      # return xml core output
      self$core <- xml_node_create(
        "cp:coreProperties",
        xml_attributes = c(
          `xmlns:cp`       = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
          `xmlns:dc`       = "http://purl.org/dc/elements/1.1/",
          `xmlns:dcterms`  = "http://purl.org/dc/terms/",
          `xmlns:dcmitype` = "http://purl.org/dc/dcmitype/",
          `xmlns:xsi`      = "http://www.w3.org/2001/XMLSchema-instance"
        ),
        xml_children = xml_properties,
        escapes = TRUE
      )

      if (!is.null(custom)) {

        if (!is.null(names(custom))) {
          custom <- mapply(
            custom, names(custom),
            FUN = function(x, y) {

              child <- xml_node_create("vt:lpwstr", xml_children = x)
              if (is.logical(x)) {
                child <- xml_node_create("vt:bool", xml_children = as_xml_attr(x))
              } else if (is.numeric(x) && is.integer(x)) {
                child <- xml_node_create("vt:i4", xml_children = as_xml_attr(x))
              } else if (is.numeric(x) && !is.integer(x)) {
                child <- xml_node_create("vt:r8", xml_children = as_xml_attr(x))
              } else if (inherits(x, "Date") || inherits(x, "POSIXt")) {
                child <- xml_node_create("vt:filetime", xml_children = format(as_POSIXct_utc(x), "%Y-%m-%dT%H:%M:%SZ"))
              }

              xml_node_create(
                "property",
                xml_attributes = c(
                  fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                  pid   = "0",
                  name  = y
                ),
                xml_children = child
              )
            },
            USE.NAMES = FALSE
          )
        }

        custom <- xml_node(custom, "property")

        if (length(self$custom) == 0) {
          self$custom <- xml_node_create(
            "Properties",
            xml_attributes = c(
              xmlns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
              `xmlns:vt` = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
            ),
            xml_children = custom
          )
          self$append(
            "Content_Types",
            "<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>"
          )
        } else {

          props <- xml_node(self$custom, "Properties", "property")

          new_names <- rbindlist(xml_attr(custom, "property"))$name
          old_names <- rbindlist(xml_attr(props, "property"))$name

          # TODO add update or remove option
          if (anyDuplicated(c(old_names, new_names))) {
            message("File has duplicated custom section")
            cstm <- self$custom
            idxs <- which(old_names %in% new_names)
            # remove all duplicates in reverse order
            for (idx in rev(idxs)) {
              cstm <- xml_rm_child(cstm, "property", which = idx)
            }
            # add replacement childs (order might differ. does it matter?)
            self$custom <- xml_add_child(cstm, custom)
          } else {
            self$custom <- xml_add_child(self$custom, custom) # pxml()
          }

        }

        self$custom <- wb_upd_custom_pid(self)
      }

      invisible(self)

    },

    #' @description add mips string
    #' @param xml A mips string added to self$custom
    add_mips = function(xml = NULL) {
      if (!is.null(xml)) assert_class(xml, "character")

      # get option and make sure that it can be imported as xml
      mips <- xml %||% getOption("openxlsx2.mips_xml_string")
      if (is.null(mips)) stop("no mips xml provided")

      nam <- xml_node_name(mips)

      if (all(nam == "clbl:labelList")) {
        self$docMetadata <- xml_node(mips, nam)
      } else {
        mips <- xml_node(mips, "property")
        self$set_properties(custom = mips)
      }

      invisible(self)
    },

    #' @description get mips string
    #' @param single_xml single_xml
    #' @param quiet quiet
    get_mips = function(single_xml = TRUE, quiet = TRUE) {
        props <- xml_node(self$custom, "Properties", "property")
        prop_nams <- grepl("MSIP_Label_", rbindlist(xml_attr(props, "property"))$name)

        name <- grepl("_Name$", rbindlist(xml_attr(props[prop_nams], "property"))$name)
        name <- xml_value(props[prop_nams][name], "property", "vt:lpwstr")
        mips <- props[prop_nams]

        if (length(name) == 0 && length(self$docMetadata)) {
          name <- xml_attr(self$docMetadata, "clbl:labelList", "clbl:label")[[1]][["id"]]
          mips <- self$docMetadata
          # names(mips) <- "docMetadata"
          single_xml <- FALSE
        }

        if (!quiet) message("Found MIPS section: ", name)

        if (single_xml)
          paste0(mips, collapse = "")
        else
          mips
    },

    #' @description Set creator(s)
    #' @param creators A character vector of creators to set.  Duplicates are
    #'   ignored.
    set_creators = function(creators) {
      self$set_properties(creator = creators)
    },

    #' @description Add creator(s)
    #' @param creators A character vector of creators to add.  Duplicates are
    #'   ignored.
    add_creators = function(creators) {
      creators <- paste0(self$get_properties()[["creator"]], ";", creators)
      self$set_properties(creator = creators)
    },

    #' @description Remove creator(s)
    #' @param creators A character vector of creators to remove.  All duplicated
    #'   are removed.
    remove_creators = function(creators) {
      old <- strsplit(self$get_properties()[["creator"]], ";")[[1]]
      old <- old[which(!old %in% creators)]
      self$set_properties(creator = old)
    },

    #' @description
    #' Change the last modified by
    #' @param name A new value
    #' @return The `wbWorkbook` object, invisibly
    set_last_modified_by = function(name, ...) {
      if (missing(name) && list(...)$LastModifiedBy) {
        .Deprecated(old = "LastModifiedBy", new = "name", package = "openxlsx2")
        name <- list(...)$LastModifiedBy
      }
      self$set_properties(modifier = name)
    },

    #' @description set_page_setup() this function is intended to supersede page_setup(), but is not yet stable
    #' @param orientation orientation
    #' @param black_and_white black_and_white
    #' @param cell_comments cell_comment
    #' @param copies copies
    #' @param draft draft
    #' @param errors errors
    #' @param first_page_number first_page_number
    #' @param id id
    #' @param page_order page_order
    #' @param paper_height,paper_width paper size
    #' @param use_first_page_number use_first_page_number
    #' @param use_printer_defaults use_printer_defaults
    #' @param hdpi,vdpi horizontal and vertical dpi
    #' @param scale scale
    #' @param left left
    #' @param right right
    #' @param top top
    #' @param bottom bottom
    #' @param header header
    #' @param footer footer
    #' @param fit_to_width fitToWidth
    #' @param fit_to_height fitToHeight
    #' @param paper_size paperSize
    #' @param print_title_rows printTitleRows
    #' @param print_title_cols printTitleCols
    #' @param summary_row summaryRow
    #' @param summary_col summaryCol
    #' @param tab_color tabColor
    #' @param ... additional arguments
    #' @return The `wbWorkbook` object, invisibly
    set_page_setup = function(
      sheet                 = current_sheet(),
      # page properties
      black_and_white       = NULL,
      cell_comments         = NULL,
      copies                = NULL,
      draft                 = NULL,
      errors                = NULL,
      first_page_number     = NULL,
      id                    = NULL, # useful and should the user be able to set this by accident?
      page_order            = NULL,
      paper_height          = NULL,
      paper_width           = NULL,
      hdpi                  = NULL,
      vdpi                  = NULL,
      use_first_page_number = NULL,
      use_printer_defaults  = NULL,
      orientation           = NULL,
      scale                 = NULL,
      left                  = 0.7,
      right                 = 0.7,
      top                   = 0.75,
      bottom                = 0.75,
      header                = 0.3,
      footer                = 0.3,
      fit_to_width          = FALSE,
      fit_to_height         = FALSE,
      paper_size            = NULL,
      # outline properties
      print_title_rows      = NULL,
      print_title_cols      = NULL,
      summary_row           = NULL,
      summary_col           = NULL,
      # tabColor properties
      tab_color             = NULL,
      ...
    ) {

      standardize(...)

      sheet <- private$get_sheet_index(sheet)
      xml <- self$worksheets[[sheet]]$pageSetup

      attrs <- rbindlist(xml_attr(xml, "pageSetup"))

      ## orientation ----
      orientation <- orientation %||% attrs$orientation
      orientation <- tolower(orientation)
      if (!orientation %in% c("portrait", "landscape")) stop("Invalid page orientation.")

      ## scale ----
      if (!is.null(scale)) {
        scale <- scale %||% attrs$scale
        scale <- as.numeric(scale)
        if ((scale < 10) || (scale > 400)) {
          message("Scale must be between 10 and 400. Scale was: ", scale)
          scale <- if (scale < 10) 10 else if (scale > 400) 400
        }
      }

      paper_size <- paper_size %||% attrs$paperSize
      if (!is.null(paper_size)) {
        paper_sizes <- 1:118
        paper_size  <- as.integer(paper_size)
        if (!paper_size %in% paper_sizes) {
          stop("paper_size must be an integer in range [1, 118]. See ?wb_page_setup details.")
        }
      }

      ## HDPI/VDPI ----
      horizontal_dpi <- hdpi %||% attrs$horizontalDpi
      vertical_dpi   <- vdpi %||% attrs$verticalDpi

      xml <- xml_attr_mod(
        xml,
        xml_attributes = c(
          blackAndWhite      = as_xml_attr(black_and_white),
          cellComments       = as_xml_attr(cell_comments),
          copies             = as_xml_attr(copies),
          draft              = as_xml_attr(draft),
          errors             = as_xml_attr(errors),
          firstPageNumber    = as_xml_attr(first_page_number),
          fitToHeight        = as_xml_attr(fit_to_height),
          fitToWidth         = as_xml_attr(fit_to_width),
          horizontalDpi      = as_xml_attr(horizontal_dpi),
          id                 = as_xml_attr(id),
          orientation        = as_xml_attr(orientation),
          pageOrder          = as_xml_attr(page_order),
          paperHeight        = as_xml_attr(paper_height),
          paperSize          = as_xml_attr(paper_size),
          paperWidth         = as_xml_attr(paper_width),
          scale              = as_xml_attr(scale),
          useFirstPageNumber = as_xml_attr(use_first_page_number),
          usePrinterDefaults = as_xml_attr(use_printer_defaults),
          verticalDpi        = as_xml_attr(vertical_dpi)
        )
      )

      self$worksheets[[sheet]]$pageSetup <- xml

      ## update pageMargins
      self$worksheets[[sheet]]$pageMargins <-
        sprintf(
          '<pageMargins left="%s" right="%s" top="%s" bottom="%s" header="%s" footer="%s"/>',
           left, right, top, bottom, header, footer
        )

      ## summary row and col ----
      outlinepr <- character()

      if (!is.null(summary_row)) {

        if (!validRow(summary_row)) {
          stop("Invalid \`summary_row\` option. Must be one of \"Above\" or \"Below\".")
        } else if (tolower(summary_row) == "above") {
          outlinepr <- c(summaryBelow = "0")
        } else {
          outlinepr <- c(summaryBelow = "1")
        }
      }

      if (!is.null(summary_col)) {

        if (!validCol(summary_col)) {
          stop("Invalid \`summary_col\` option. Must be one of \"Left\" or \"Right\".")
        } else if (tolower(summary_col) == "left") {
          outlinepr <- c(outlinepr, c(summaryRight = "0"))
        } else {
          outlinepr <- c(outlinepr, c(summaryRight = "1"))
        }
      }

      ## update sheetPr ----
      xml <- self$worksheets[[sheet]]$sheetPr

      if (length(xml) == 0) xml <- "<sheetPr/>"

      sheetpr_df <- read_sheetpr(xml)

      ## order matters: tabColor, outlinePr, pageSetUpPr.
      if (length(tab_color)) {
        tc <- sheetpr_df$tabColor
        if (tc == "") tc <- "<tabColor/>"
        if (is.null(names(tab_color))) {
          if (tab_color == "") {
            tab_color <- NULL
          } else {
            warning("tab_color should be a wb_color() object")
            tab_color <- wb_color(tab_color)
          }
        }

        if (is.null(tab_color)) {
          sheetpr_df$tabColor <- ""
        } else {
          sheetpr_df$tabColor <- xml_attr_mod(tc, xml_attributes = tab_color)
        }
      }

      ## TODO make sure that the order is valid
      if (length(outlinepr)) {
        op <- sheetpr_df$outlinePr
        if (op == "") op <- "<outlinePr/>"
        sheetpr_df$outlinePr <- xml_attr_mod(op, xml_attributes = outlinepr)
      }

      if (fit_to_height || fit_to_width) {
        psup <- sheetpr_df$pageSetUpPr
        if (psup == "") psup <- "<pageSetUpPr/>"
        sheetpr_df$pageSetUpPr <- xml_attr_mod(psup, xml_attributes = c(fitToPage = "1"))
      }

      self$worksheets[[sheet]]$sheetPr <- write_sheetpr(sheetpr_df)

      ## print Titles ----
      if (!is.null(print_title_rows) && is.null(print_title_cols)) {
        if (!is.numeric(print_title_rows)) {
          stop("print_title_rows must be numeric.")
        }

        private$create_named_region(
          ref1 = paste0("$", min(print_title_rows)),
          ref2 = paste0("$", max(print_title_rows)),
          name = "_xlnm.Print_Titles",
          sheet = self$get_sheet_names(escape = TRUE)[[sheet]],
          localSheetId = sheet - 1L
        )
      } else if (!is.null(print_title_cols) && is.null(print_title_rows)) {
        if (!is.numeric(print_title_cols)) {
          stop("print_title_cols must be numeric.")
        }

        cols <- int2col(range(print_title_cols))
        private$create_named_region(
          ref1 = paste0("$", cols[1]),
          ref2 = paste0("$", cols[2]),
          name = "_xlnm.Print_Titles",
          sheet = self$get_sheet_names(escape = TRUE)[[sheet]],
          localSheetId = sheet - 1L
        )
      } else if (!is.null(print_title_cols) && !is.null(print_title_rows)) {
        if (!is.numeric(print_title_rows)) {
          stop("print_title_rows must be numeric.")
        }

        if (!is.numeric(print_title_cols)) {
          stop("print_title_cols must be numeric.")
        }

        cols <- int2col(range(print_title_cols))
        rows <- range(print_title_rows)

        cols <- paste(paste0("$", cols[1]), paste0("$", cols[2]), sep = ":")
        rows <- paste(paste0("$", rows[1]), paste0("$", rows[2]), sep = ":")
        localSheetId <- sheet - 1L
        sheet <- self$get_sheet_names(escape = TRUE)[[sheet]]

        self$workbook$definedNames <- c(
          self$workbook$definedNames,
          sprintf(
            '<definedName name="_xlnm.Print_Titles" localSheetId="%s">\'%s\'!%s,\'%s\'!%s</definedName>',
            localSheetId, sheet, cols, sheet, rows
          )
        )

      }

      invisible(self)
    },

    #' @description page_setup()
    #' @param orientation orientation
    #' @param scale scale
    #' @param left left
    #' @param right right
    #' @param top top
    #' @param bottom bottom
    #' @param header header
    #' @param footer footer
    #' @param fit_to_width fitToWidth
    #' @param fit_to_height fitToHeight
    #' @param paper_size paperSize
    #' @param print_title_rows printTitleRows
    #' @param print_title_cols printTitleCols
    #' @param summary_row summaryRow
    #' @param summary_col summaryCol
    #' @return The `wbWorkbook` object, invisibly
    page_setup = function(
      sheet            = current_sheet(),
      orientation      = NULL,
      scale            = 100,
      left             = 0.7,
      right            = 0.7,
      top              = 0.75,
      bottom           = 0.75,
      header           = 0.3,
      footer           = 0.3,
      fit_to_width     = FALSE,
      fit_to_height    = FALSE,
      paper_size       = NULL,
      print_title_rows = NULL,
      print_title_cols = NULL,
      summary_row      = NULL,
      summary_col      = NULL,
      ...
    ) {

      standardize_case_names(...)

      sheet <- private$get_sheet_index(sheet)

      self$set_page_setup(
        sheet            = sheet,
        orientation      = orientation,
        scale            = scale,
        left             = left,
        right            = right,
        top              = top,
        bottom           = bottom,
        header           = header,
        footer           = footer,
        fit_to_width     = fit_to_width,
        fit_to_height    = fit_to_height,
        paper_size       = paper_size,
        print_title_rows = print_title_rows,
        print_title_cols = print_title_cols,
        summary_row      = summary_row,
        summary_col      = summary_col
      )

      invisible(self)
    },

    ## header footer ----

    #' @description Sets headers and footers
    #' @param header header
    #' @param footer footer
    #' @param even_header evenHeader
    #' @param even_footer evenFooter
    #' @param first_header firstHeader
    #' @param first_footer firstFooter
    #' @param align_with_margins align_with_margins
    #' @param scale_with_doc scale_with_doc
    #' @return The `wbWorkbook` object, invisibly
    set_header_footer = function(
      sheet              = current_sheet(),
      header             = NULL,
      footer             = NULL,
      even_header        = NULL,
      even_footer        = NULL,
      first_header       = NULL,
      first_footer       = NULL,
      align_with_margins = NULL,
      scale_with_doc     = NULL,
      ...
    ) {

      standardize_case_names(...)

      sheet <- private$get_sheet_index(sheet)

      not_three_or_na <- function(x) {
        nam <- deparse(substitute(x))
        msg <- sprintf(
          "`%s` must have length 3 where elements correspond to positions: left, center, right.",
          nam
        )

        if (!is.null(x) && !(length(x) == 3 || (length(x) == 1 && is.na(x))))
          stop(msg, call. = FALSE)
      }

      not_three_or_na(header)
      not_three_or_na(footer)
      not_three_or_na(even_header)
      not_three_or_na(even_footer)
      not_three_or_na(first_header)
      not_three_or_na(first_footer)

      # TODO this could probably be moved to the hf assignment
      oddHeader   <- headerFooterSub(header)
      oddFooter   <- headerFooterSub(footer)
      evenHeader  <- headerFooterSub(even_header)
      evenFooter  <- headerFooterSub(even_footer)
      firstHeader <- headerFooterSub(first_header)
      firstFooter <- headerFooterSub(first_footer)

      hf <- list(
        oddHeader   = naToNULLList(oddHeader),
        oddFooter   = naToNULLList(oddFooter),
        evenHeader  = naToNULLList(evenHeader),
        evenFooter  = naToNULLList(evenFooter),
        firstHeader = naToNULLList(firstHeader),
        firstFooter = naToNULLList(firstFooter)
      )

      if (all(lengths(hf) == 0)) {
        hf <- NULL
      } else {
        if (!is.null(old_hf <- self$worksheets[[sheet]]$headerFooter)) {
          for (nam in names(hf)) {
            # Update using new_vector if it exists, keeping original values where NA
            if (length(hf[[nam]]) && length(old_hf[[nam]])) {
              sel <- if (is.list(hf[[nam]]) && length(hf[[nam]]) == 0) sel <- seq_len(3)
                      else which(vapply(hf[[nam]], is.null, NA))
              hf[[nam]][sel] <- old_hf[[nam]][sel]
            }
          }
        }
      }

      if (!is.null(scale_with_doc)) {
        assert_class(scale_with_doc, "logical")
        self$worksheets[[sheet]]$scale_with_doc     <- scale_with_doc
      }

      if (!is.null(align_with_margins)) {
        assert_class(align_with_margins, "logical")
        self$worksheets[[sheet]]$align_with_margins <- align_with_margins
      }

      self$worksheets[[sheet]]$headerFooter <- hf
      invisible(self)
    },

    #' @description get tables
    #' @return The sheet tables.  `character()` if empty
    get_tables = function(sheet = current_sheet()) {
      if (!is.null(sheet) && length(sheet) != 1) {
        stop("sheet argument must be length 1")
      }

      if (is.null(self$tables)) {
        return(character())
      }

      if (!is.null(sheet)) {
        sheet <- private$get_sheet_index(sheet)
        if (is.na(sheet)) stop("No such sheet in workbook")

        sel <- self$tables$tab_sheet == sheet & self$tables$tab_act == 1
      } else {
        sel <- self$tables$tab_act == 1
      }
      self$tables[sel, c("tab_name", "tab_ref")]
    },


    #' @description remove tables
    #' @param table table
    #' @param remove_data removes the data as well
    #' @return The `wbWorkbook` object
    remove_tables = function(sheet = current_sheet(), table, remove_data = TRUE) {
      if (length(table) != 1) {
        stop("table argument must be length 1")
      }

      ## delete table object and all data in it
      sheet <- private$get_sheet_index(sheet)

      if (!table %in% self$tables$tab_name) {
        stop(sprintf("table '%s' does not exist.\n
                     Call `wb_get_tables()` to get existing table names", table),
             call. = FALSE)
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

      ## now delete data
      if (remove_data)
        self$clean_sheet(sheet = sheet, dims = refs)

      invisible(self)
    },

    #' @description add filters
    #' @param rows rows
    #' @param cols cols
    #' @return The `wbWorkbook` object
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
        paste(get_cell_refs(data.frame("x" = c(rows, rows), "y" = c(min(cols), max(cols)), stringsAsFactors = FALSE)), collapse = ":")
      )

      invisible(self)
    },

    #' @description remove filters
    #' @return The `wbWorkbook` object
    remove_filter = function(sheet = current_sheet()) {
      for (s in private$get_sheet_index(sheet)) {
        self$worksheets[[s]]$autoFilter <- character()
      }

      invisible(self)
    },

    #' @description grid lines
    #' @param show show
    #' @param print print
    #' @return The `wbWorkbook` object
    set_grid_lines = function(sheet = current_sheet(), show = FALSE, print = show) {
      sheet <- private$get_sheet_index(sheet)

      assert_class(show, "logical")
      assert_class(print, "logical")

      ## show
      self$worksheets[[sheet]]$set_sheetview(show_grid_lines = as_xml_attr(show))

      ## print
      self$worksheets[[sheet]]$set_print_options(gridLines = as_xml_attr(print), gridLinesSet = as_xml_attr(print))

      invisible(self)
    },

    #' @description grid lines
    #' @param show show
    #' @param print print
    #' @return The `wbWorkbook` object
    grid_lines = function(sheet = current_sheet(), show = FALSE, print = show) {
      .Deprecated(old = "grid_lines", new = "set_grid_lines", package = "openxlsx2")
      self$set_grid_lines(sheet = sheet, show = show, print = print)
    },

    ### named region ----

    #' @description add a named region
    #' @param name name
    #' @param local_sheet local_sheet
    #' @param overwrite overwrite
    #' @param comment comment
    #' @param custom_menu custom_menu
    #' @param description description
    #' @param is_function function
    #' @param function_group_id function group id
    #' @param help help
    #' @param hidden hidden
    #' @param local_name localName
    #' @param publish_to_server publish to server
    #' @param status_bar status bar
    #' @param vb_procedure vb procedure
    #' @param workbook_parameter workbookParameter
    #' @param xml xml
    #' @return The `wbWorkbook` object
    add_named_region = function(
      sheet = current_sheet(),
      dims = "A1",
      name,
      local_sheet        = FALSE,
      overwrite          = FALSE,
      comment            = NULL,
      hidden             = NULL,
      custom_menu        = NULL,
      description        = NULL,
      is_function        = NULL,
      function_group_id  = NULL,
      help               = NULL,
      local_name         = NULL,
      publish_to_server  = NULL,
      status_bar         = NULL,
      vb_procedure       = NULL,
      workbook_parameter = NULL,
      xml                = NULL,
      ...
    ) {

      arguments <- c(ls(), "rows", "cols")
      standardize_case_names(..., arguments = arguments)

      sheet <- private$get_sheet_index(sheet)

      cols <- list(...)[["cols"]]
      rows <- list(...)[["rows"]]

      if (!is.null(rows) && !is.null(cols)) {
        .Deprecated(old = "cols/rows", new = "dims", package = "openxlsx2")
        dims <- rowcol_to_dims(rows, cols)
      }

      localSheetId <- ""
      if (local_sheet) localSheetId <- as.character(sheet - 1L)

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

      assert_named_region(name)

      if (any(match_dn)) {
        if (overwrite)
          self$workbook$definedNames <- self$workbook$definedNames[-match_dn]
        else
          stop(sprintf("Named region with name '%s' already exists! Use overwrite  = TRUE if you want to replace it", name))
      }

      rowcols <- dims_to_rowcol(dims, as_integer = TRUE)
      rows <- rowcols[["row"]]
      cols <- rowcols[["col"]]

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
        customMenu         = custom_menu,
        description        = description,
        is_function        = is_function,
        functionGroupId    = function_group_id,
        help               = help,
        hidden             = hidden,
        localName          = local_name,
        publishToServer    = publish_to_server,
        statusBar          = status_bar,
        vbProcedure        = vb_procedure,
        workbookParameter  = workbook_parameter,
        xml                = xml
      )

      invisible(self)
    },

    #' @description get named regions in a workbook
    #' @param tables Return tables as well?
    #' @param x Not used.
    #' @return A `data.frame` of named regions
    get_named_regions = function(tables = FALSE, x = NULL) {
      if (!is.null(x)) {
        stop("x should not be provided to get_named_regions.", call. = FALSE)
      }
      z <- NULL

      if (length(self$workbook$definedNames)) {
        z <- get_nr_from_definedName(self)
      }

      if (tables && !is.null(self$tables)) {
        tb <- get_named_regions_tab(self)

        if (is.null(z)) {
          z <- tb
        } else {
          z <- merge(z, tb, all = TRUE, sort = FALSE)
        }

      }

      z
    },
    #' @description remove a named region
    #' @param name name
    #' @return The `wbWorkbook` object
    remove_named_region = function(sheet = current_sheet(), name = NULL) {
      # get all nown defined names
      dn <- wb_get_named_regions(self)

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
    #' @return Returns sheet visibility
    get_sheet_visibility = function() {
      state <- rep("visible", length(self$workbook$sheets))
      state[grepl("hidden", self$workbook$sheets)] <- "hidden"
      state[grepl("veryHidden", self$workbook$sheets, ignore.case = TRUE)] <- "veryHidden"
      state
    },

    #' @description Set sheet visibility
    #' @param value value
    #' @return The `wbWorkbook` object
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
    #' @param row row
    #' @param col col
    #' @return The `wbWorkbook` object
    add_page_break = function(sheet = current_sheet(), row = NULL, col = NULL) {
      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$add_page_break(row = row, col = col)
      invisible(self)
    },

    #' @description clean sheet (remove all values)
    #' @param numbers remove all numbers
    #' @param characters remove all characters
    #' @param styles remove all styles
    #' @param merged_cells remove all merged_cells
    #' @param hyperlinks remove all hyperlinks
    #' @return The `wbWorksheetObject`, invisibly
    clean_sheet = function(
        sheet        = current_sheet(),
        dims         = NULL,
        numbers      = TRUE,
        characters   = TRUE,
        styles       = TRUE,
        merged_cells = TRUE,
        hyperlinks   = TRUE
    ) {
      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$clean_sheet(
        dims         = dims,
        numbers      = numbers,
        characters   = characters,
        styles       = styles,
        merged_cells = merged_cells
      )

      if (hyperlinks)
        self$remove_hyperlink(
          sheet = sheet,
          dims  = dims
        )

      invisible(self)
    },

    #' @description create borders for cell region
    #' @param dims dimensions on the worksheet e.g. "A1", "A1:A5", "A1:H5"
    #' @param bottom_color,left_color,right_color,top_color,inner_hcolor,inner_vcolor a color, either something openxml knows or some RGB color
    #' @param left_border,right_border,top_border,bottom_border,inner_hgrid,inner_vgrid the border style, if NULL no border is drawn. See create_border for possible border styles
    #' @param update update
    #' @return The `wbWorkbook`, invisibly
    add_border = function(
      sheet         = current_sheet(),
      dims          = "A1",
      bottom_color  = wb_color(hex = "FF000000"),
      left_color    = wb_color(hex = "FF000000"),
      right_color   = wb_color(hex = "FF000000"),
      top_color     = wb_color(hex = "FF000000"),
      bottom_border = "thin",
      left_border   = "thin",
      right_border  = "thin",
      top_border    = "thin",
      inner_hgrid   = NULL,
      inner_hcolor  = NULL,
      inner_vgrid   = NULL,
      inner_vcolor  = NULL,
      update        = FALSE,
      ...
    ) {

      # TODO merge styles and if a style is already present, only add the newly
      # created border style

      # cc <- wb$worksheets[[sheet]]$sheet_data$cc
      # df_s <- as.data.frame(lapply(df, function(x) cc$c_s[cc$r %in% x]))

      standardize(...)

      if (is.null(bottom_color)) bottom_border <- NULL
      if (is.null(left_color)) left_border <- NULL
      if (is.null(right_color)) right_border <- NULL
      if (is.null(top_color)) top_border <- NULL

      if (is.null(bottom_border)) bottom_color <- NULL
      if (is.null(left_border)) left_color <- NULL
      if (is.null(right_border)) right_color <- NULL
      if (is.null(top_border)) top_color <- NULL

      df <- dims_to_dataframe(dims, fill = TRUE)
      sheet <- private$get_sheet_index(sheet)

      private$do_cell_init(sheet, dims)

      ### border creation

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

      smp <- random_string()
      ### full-single
      if (ncol(df) == 1 && nrow(df) == 1) {
        # create border
        full_single <- create_border(
          top = top_border, top_color = top_color,
          bottom = bottom_border, bottom_color = bottom_color,
          left = left_border, left_color = left_color,
          right = right_border, right_color = right_color
        )

        # determine dim
        dim_full_single <- df[1, 1]

        if (update) full_single <- update_border(self, dims = dim_full_single, new_border = full_single)

        # determine name
        sfull_single <- paste0(smp, "full_single")

        # add border
        self$styles_mgr$add(full_single, sfull_single)
        xf_prev <- get_cell_styles(self, sheet, dims)
        xf_full_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sfull_single))
        self$styles_mgr$add(xf_full_single, xf_full_single)
        self$set_cell_style(sheet, dims, self$styles_mgr$get_xf_id(xf_full_single))
      }

      ### single
      if (ncol(df) == 1 && nrow(df) > 1) {
        # create borders
        top_single <- create_border(
          top = top_border, top_color = top_color,
          bottom = inner_hgrid, bottom_color = inner_hcolor,
          left = left_border, left_color = left_color,
          right = right_border, right_color = right_color
        )

        bottom_single <- create_border(
          top = inner_hgrid, top_color = inner_hcolor,
          bottom = bottom_border, bottom_color = bottom_color,
          left = left_border, left_color = left_color,
          right = right_border, right_color = right_color
        )

        # determine dims
        dim_top_single <- df[1, 1]
        dim_bottom_single <- df[nrow(df), 1]

        if (update) {
          top_single    <- update_border(self, dims = dim_top_single, new_border = top_single)
          bottom_single <- update_border(self, dims = dim_bottom_single, new_border = bottom_single)
        }

        # determine names
        stop_single <- paste0(smp, "full_single")
        sbottom_single <- paste0(smp, "bottom_single")

        # add top single
        self$styles_mgr$add(top_single, stop_single)
        xf_prev <- get_cell_styles(self, sheet, dim_top_single)
        xf_top_single <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_single))
        self$styles_mgr$add(xf_top_single, xf_top_single)
        self$set_cell_style(sheet, dim_top_single, self$styles_mgr$get_xf_id(xf_top_single))

        # add bottom single
        self$styles_mgr$add(bottom_single, sbottom_single)
        xf_prev <- get_cell_styles(self, sheet, dim_bottom_single)
        xf_bottom_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_single))
        self$styles_mgr$add(xf_bottom_single, xf_bottom_single)
        self$set_cell_style(sheet, dim_bottom_single, self$styles_mgr$get_xf_id(xf_bottom_single))

        if (nrow(df) >= 3) {
          # create border
          middle_single <- create_border(
            top = inner_hgrid, top_color = inner_hcolor,
            bottom = inner_hgrid, bottom_color = inner_hcolor,
            left = left_border, left_color = left_color,
            right = right_border, right_color = right_color
          )

          # determine dims
          mid <- df[, 1]
          dim_middle_single <- mid[!mid %in% c(dim_top_single, dim_bottom_single)]

          if (update) middle_single <- update_border(self, dims = dim_middle_single, new_border = middle_single)

          # determine names
          smiddle_single <- paste0(smp, "middle_single", seq_along(middle_single))

          # add middle single
          self$styles_mgr$add(middle_single, smiddle_single)
          xf_prev <- get_cell_styles(self, sheet, dim_middle_single)
          xf_middle_single <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_single))
          self$styles_mgr$add(xf_middle_single, xf_middle_single)
          self$set_cell_style(sheet, dim_middle_single, self$styles_mgr$get_xf_id(xf_middle_single))
        }
      }

      # create left and right single row pieces
      if (ncol(df) >= 2 && nrow(df) == 1) {
        # create borders
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

        # determine dims
        dim_left_single <- df[1, 1]
        dim_right_single <- df[1, ncol(df)]

        if (update) {
          left_single  <- update_border(self, dims = dim_left_single, new_border = left_single)
          right_single <- update_border(self, dims = dim_right_single, new_border = right_single)
        }

        # determine names
        sleft_single <- paste0(smp, "left_single")
        sright_single <- paste0(smp, "right_single")

        # add left single
        self$styles_mgr$add(left_single, sleft_single)
        xf_prev <- get_cell_styles(self, sheet, dim_left_single)
        xf_left_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sleft_single))
        self$styles_mgr$add(xf_left_single, xf_left_single)
        self$set_cell_style(sheet, dim_left_single, self$styles_mgr$get_xf_id(xf_left_single))

        # add right single
        self$styles_mgr$add(right_single, sright_single)
        xf_prev <- get_cell_styles(self, sheet, dim_right_single)
        xf_right_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sright_single))
        self$styles_mgr$add(xf_right_single, xf_right_single)
        self$set_cell_style(sheet, dim_right_single, self$styles_mgr$get_xf_id(xf_right_single))

        # add single center piece(s)
        if (ncol(df) >= 3) {
          center_single <- create_border(
            top = top_border, top_color = top_color,
            bottom = bottom_border, bottom_color = bottom_color,
            left = inner_vgrid, left_color = inner_vcolor,
            right = inner_vgrid, right_color = inner_vcolor
          )

          # determine dims
          ctr <- df[1, ]
          dim_center_single <- ctr[!ctr %in% c(dim_left_single, dim_right_single)]

          if (update) center_single <- update_border(self, dims = dim_center_single, new_border = center_single)

          # determine names
          scenter_single <- paste0(smp, "center_single", seq_along(center_single))

          # add center single
          self$styles_mgr$add(center_single, scenter_single)
          xf_prev <- get_cell_styles(self, sheet, dim_center_single)
          xf_center_single <- set_border(xf_prev, self$styles_mgr$get_border_id(scenter_single))
          self$styles_mgr$add(xf_center_single, xf_center_single)
          self$set_cell_style(sheet, dim_center_single, self$styles_mgr$get_xf_id(xf_center_single))
        }

      }

      # create left & right - top & bottom corners pieces
      if (ncol(df) >= 2 && nrow(df) >= 2) {
        # create borders
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
          right = right_border, right_color = right_color
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
          right = right_border, right_color = right_color
        )

        # determine dims
        dim_top_left     <- df[1, 1]
        dim_bottom_left  <- df[nrow(df), 1]
        dim_top_right    <- df[1, ncol(df)]
        dim_bottom_right <- df[nrow(df), ncol(df)]

        if (update) {
          top_left     <- update_border(self, dims = dim_top_left, new_border = top_left)
          bottom_left  <- update_border(self, dims = dim_bottom_left, new_border = bottom_left)
          top_right    <- update_border(self, dims = dim_top_right, new_border = top_right)
          bottom_right <- update_border(self, dims = dim_bottom_right, new_border = bottom_right)
        }

        # determine names
        stop_left <- paste0(smp, "top_left")
        sbottom_left <- paste0(smp, "bottom_left")
        stop_right <- paste0(smp, "top_right")
        sbottom_right <- paste0(smp, "bottom_right")

        # add top left
        self$styles_mgr$add(top_left, stop_left)
        xf_prev <- get_cell_styles(self, sheet, dim_top_left)
        xf_top_left <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_left))
        self$styles_mgr$add(xf_top_left, xf_top_left)
        self$set_cell_style(sheet, dim_top_left, self$styles_mgr$get_xf_id(xf_top_left))

        # add top right
        self$styles_mgr$add(top_right, stop_right)
        xf_prev <- get_cell_styles(self, sheet, dim_top_right)
        xf_top_right <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_right))
        self$styles_mgr$add(xf_top_right, xf_top_right)
        self$set_cell_style(sheet, dim_top_right, self$styles_mgr$get_xf_id(xf_top_right))

        # add bottom left
        self$styles_mgr$add(bottom_left, sbottom_left)
        xf_prev <- get_cell_styles(self, sheet, dim_bottom_left)
        xf_bottom_left <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_left))
        self$styles_mgr$add(xf_bottom_left, xf_bottom_left)
        self$set_cell_style(sheet, dim_bottom_left, self$styles_mgr$get_xf_id(xf_bottom_left))

        # add bottom right
        self$styles_mgr$add(bottom_right, sbottom_right)
        xf_prev <- get_cell_styles(self, sheet, dim_bottom_right)
        xf_bottom_right <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_right))
        self$styles_mgr$add(xf_bottom_right, xf_bottom_right)
        self$set_cell_style(sheet, dim_bottom_right, self$styles_mgr$get_xf_id(xf_bottom_right))
      }

      # create left and right middle pieces
      if (ncol(df) >= 2 && nrow(df) >= 3) {
        # create borders
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

        # determine dims
        top_mid <- df[, 1]
        bottom_mid <- df[, ncol(df)]
        dim_middle_left <- top_mid[!top_mid %in% c(dim_top_left, dim_bottom_left)]
        dim_middle_right <- bottom_mid[!bottom_mid %in% c(dim_top_right, dim_bottom_right)]

        if (update) {
          middle_left  <- update_border(self, dims = dim_middle_left, new_border = middle_left)
          middle_right <- update_border(self, dims = dim_middle_right, new_border = middle_right)
        }

        # determine names
        smiddle_left <- paste0(smp, "middle_left", seq_along(middle_left))
        smiddle_right <- paste0(smp, "middle_right", seq_along(middle_right))

        # add middle left
        self$styles_mgr$add(middle_left, smiddle_left)
        xf_prev <- get_cell_styles(self, sheet, dim_middle_left)
        xf_middle_left <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_left))
        self$styles_mgr$add(xf_middle_left, xf_middle_left)
        self$set_cell_style(sheet, dim_middle_left, self$styles_mgr$get_xf_id(xf_middle_left))

        # add middle right
        self$styles_mgr$add(middle_right, smiddle_right)
        xf_prev <- get_cell_styles(self, sheet, dim_middle_right)
        xf_middle_right <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_right))
        self$styles_mgr$add(xf_middle_right, xf_middle_right)
        self$set_cell_style(sheet, dim_middle_right, self$styles_mgr$get_xf_id(xf_middle_right))
      }

      # create top and bottom center pieces
      if (ncol(df) >= 3 & nrow(df) >= 2) {
        # create borders
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

        # determine dims
        top_ctr <- df[1, ]
        bottom_ctr <- df[nrow(df), ]
        dim_top_center <- top_ctr[!top_ctr %in% c(dim_top_left, dim_top_right)]
        dim_bottom_center <- bottom_ctr[!bottom_ctr %in% c(dim_bottom_left, dim_bottom_right)]

        if (update) {
          top_center    <- update_border(self, dims = dim_top_center, new_border = top_center)
          bottom_center <- update_border(self, dims = dim_bottom_center, new_border = bottom_center)
        }

        # determine names
        stop_center <- paste0(smp, "top_center", seq_along(top_center))
        sbottom_center <- paste0(smp, "bottom_center", seq_along(bottom_center))

        # add top center
        self$styles_mgr$add(top_center, stop_center)
        xf_prev <- get_cell_styles(self, sheet, dim_top_center)
        xf_top_center <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_center))
        self$styles_mgr$add(xf_top_center, xf_top_center)
        self$set_cell_style(sheet, dim_top_center, self$styles_mgr$get_xf_id(xf_top_center))

        # add bottom center
        self$styles_mgr$add(bottom_center, sbottom_center)
        xf_prev <- get_cell_styles(self, sheet, dim_bottom_center)
        xf_bottom_center <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_center))
        self$styles_mgr$add(xf_bottom_center, xf_bottom_center)
        self$set_cell_style(sheet, dim_bottom_center, self$styles_mgr$get_xf_id(xf_bottom_center))
      }

      if (nrow(df) > 2 && ncol(df) > 2) {
        # create border
        inner_cell <- create_border(
          top = inner_hgrid, top_color = inner_hcolor,
          bottom = inner_hgrid, bottom_color = inner_hcolor,
          left = inner_vgrid, left_color = inner_vcolor,
          right = inner_vgrid, right_color = inner_vcolor
        )

        # determine dims
        t_row <- 1
        b_row <- nrow(df)
        l_row <- 1
        r_row <- ncol(df)
        dim_inner_cell <- as.character(unlist(df[c(-t_row, -b_row), c(-l_row, -r_row)]))

        if (update) inner_cell <- update_border(self, dims = dim_inner_cell, new_border = inner_cell)

        # determine name
        sinner_cell <- paste0(smp, "inner_cell", seq_along(inner_cell))

        # add inner cells
        self$styles_mgr$add(inner_cell, sinner_cell)
        xf_prev <- get_cell_styles(self, sheet, dim_inner_cell)
        xf_inner_cell <- set_border(xf_prev, self$styles_mgr$get_border_id(sinner_cell))
        self$styles_mgr$add(xf_inner_cell, xf_inner_cell)
        self$set_cell_style(sheet, dim_inner_cell, self$styles_mgr$get_xf_id(xf_inner_cell))
      }

      invisible(self)
    },

    #' @description provide simple fill function
    #' @param color the colors to apply, e.g. yellow: wb_color(hex = "FFFFFF00")
    #' @param pattern various default "none" but others are possible:
    #'  "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal",
    #'  "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis",
    #'  "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid",
    #'  "lightTrellis", "gray125", "gray0625"
    #' @param gradient_fill a gradient fill xml pattern.
    #' @param every_nth_col which col should be filled
    #' @param every_nth_row which row should be filled
    #' @return The `wbWorksheetObject`, invisibly
    add_fill = function(
        sheet         = current_sheet(),
        dims          = "A1",
        color         = wb_color(hex = "FFFFFF00"),
        pattern       = "solid",
        gradient_fill = "",
        every_nth_col = 1,
        every_nth_row = 1,
        ...
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

      standardize(...)

      for (style in styles) {
        dim <- cc[cc$c_s == style, "r"]

        new_fill <- create_fill(
          gradient_fill = gradient_fill,
          pattern_type = pattern,
          fg_color = color
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
    #' @param name font name: default "Aptos Narrow"
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
    #' @param vert_align vertical alignment
    #' @param update update
    #' @return The `wbWorkbook`, invisibly
    add_font = function(
        sheet      = current_sheet(),
        dims       = "A1",
        name       = "Aptos Narrow",
        color      = wb_color(hex = "FF000000"),
        size       = "11",
        bold       = "",
        italic     = "",
        outline    = "",
        strike     = "",
        underline  = "",
        # fine tuning
        charset    = "",
        condense   = "",
        extend     = "",
        family     = "",
        scheme     = "",
        shadow     = "",
        vert_align = "",
        update     = FALSE,
        ...
    ) {
      sheet <- private$get_sheet_index(sheet)
      private$do_cell_init(sheet, dims)

      did <- dims_to_dataframe(dims, fill = TRUE)
      dims <- unname(unlist(did))

      cc <- self$worksheets[[sheet]]$sheet_data$cc
      cc <- cc[cc$r %in% dims, ]
      styles <- unique(cc[["c_s"]])

      standardize(...)

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
          vert_align = vert_align
        )

        xf_prev <- get_cell_styles(self, sheet, dim[[1]])

        if (is.character(update) || (is.logical(update) && isTRUE(update))) {
          valid <- c(
            "name", "color", "colour", "size", "bold", "italic", "outline", "strike",
            "underline", "charset", "condense", "extend", "family", "scheme", "shadow",
            "vert_align"
          )
          # update == TRUE: the user wants everything updated
          if (is.logical(update) && isTRUE(update)) {
           update <- valid[-which(valid == "colour")]
          }
          match.arg(update, valid, several.ok = TRUE)

          font_properties <- c(
            bold = "b",
            charset = "charset",
            color = "color",
            condense = "condense",
            extend = "extend",
            family = "family",
            italic = "i",
            name = "name",
            outline = "outline",
            scheme = "scheme",
            shadow = "shadow",
            strike = "strike",
            size = "sz",
            underline = "u",
            vert_align = "vertAlign"
          )
          sel <- font_properties[update]

          font_id  <- as.integer(vapply(xml_attr(xf_prev, "xf"), "[[", "fontId", FUN.VALUE = NA_character_)) + 1L
          font_xml <- self$styles_mgr$styles$fonts[[font_id]]

          # read as data frame with xml elements
          old_font <- read_font(read_xml(font_xml))
          new_font <- read_font(read_xml(new_font))

          # update elements
          old_font[sel] <- new_font[sel]

          # write as xml font
          new_font <- write_font(old_font)
        }

        self$styles_mgr$add(new_font, new_font)

        xf_new_font <- set_font(xf_prev, self$styles_mgr$get_font_id(new_font))

        self$styles_mgr$add(xf_new_font, xf_new_font)
        s_id <- self$styles_mgr$get_xf_id(xf_new_font)
        self$set_cell_style(sheet, dim, s_id)
      }

      invisible(self)
    },

    #' @description provide simple number format function
    #' @param numfmt number format id or a character of the format
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
    #' @param ext_lst extension list something like `<extLst>...</extLst>`
    #' @param hidden logical cell is hidden
    #' @param horizontal align content horizontal ('left', 'center', 'right')
    #' @param indent logical indent content
    #' @param justify_last_line logical justify last line
    #' @param locked logical cell is locked
    #' @param pivot_button unknown
    #' @param quote_prefix unknown
    #' @param reading_order reading order left to right
    #' @param relative_indent relative indentation
    #' @param shrink_to_fit logical shrink to fit
    #' @param text_rotation degrees of text rotation
    #' @param vertical vertical alignment of content ('top', 'center', 'bottom')
    #' @param wrap_text wrap text in cell
    # alignments
    #' @param apply_alignment logical apply alignment
    #' @param apply_border logical apply border
    #' @param apply_fill logical apply fill
    #' @param apply_font logical apply font
    #' @param apply_number_format logical apply number format
    #' @param apply_protection logical apply protection
    # ids
    #' @param border_id border ID to apply
    #' @param fill_id fill ID to apply
    #' @param font_id font ID to apply
    #' @param num_fmt_id number format ID to apply
    #' @param xf_id xf ID to apply
    #' @return The `wbWorkbook` object, invisibly
    add_cell_style = function(
        sheet               = current_sheet(),
        dims                = "A1",
        apply_alignment     = NULL,
        apply_border        = NULL,
        apply_fill          = NULL,
        apply_font          = NULL,
        apply_number_format = NULL,
        apply_protection    = NULL,
        border_id           = NULL,
        ext_lst             = NULL,
        fill_id             = NULL,
        font_id             = NULL,
        hidden              = NULL,
        horizontal          = NULL,
        indent              = NULL,
        justify_last_line   = NULL,
        locked              = NULL,
        num_fmt_id          = NULL,
        pivot_button        = NULL,
        quote_prefix        = NULL,
        reading_order       = NULL,
        relative_indent     = NULL,
        shrink_to_fit       = NULL,
        text_rotation       = NULL,
        vertical            = NULL,
        wrap_text           = NULL,
        xf_id               = NULL,
        ...
    ) {

      standardize_case_names(...)

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
          applyAlignment    = apply_alignment,
          applyBorder       = apply_border,
          applyFill         = apply_fill,
          applyFont         = apply_font,
          applyNumberFormat = apply_number_format,
          applyProtection   = apply_protection,
          borderId          = border_id,
          extLst            = ext_lst,
          fillId            = fill_id,
          fontId            = font_id,
          hidden            = hidden,
          horizontal        = horizontal,
          indent            = indent,
          justifyLastLine   = justify_last_line,
          locked            = locked,
          numFmtId          = num_fmt_id,
          pivotButton       = pivot_button,
          quotePrefix       = quote_prefix,
          readingOrder      = reading_order,
          relativeIndent    = relative_indent,
          shrinkToFit       = shrink_to_fit,
          textRotation      = text_rotation,
          vertical          = vertical,
          wrapText          = wrap_text,
          xfId              = xf_id
        )
        self$styles_mgr$add(xf_new_cellstyle, xf_new_cellstyle)
        s_id <- self$styles_mgr$get_xf_id(xf_new_cellstyle)
        self$set_cell_style(sheet, dim, s_id)
      }

      invisible(self)
    },

    #' @description get sheet style
    #' @return a character vector of cell styles
    get_cell_style = function(sheet = current_sheet(), dims) {

      if (length(dims) == 1 && grepl(":", dims))
        dims <- dims_to_dataframe(dims, fill = TRUE)
      sheet <- private$get_sheet_index(sheet)

      # We need to return a cell style, even if the cell is not part of the
      # workbook. Since we need to return the values in the corret order, we
      # initiate a cell, if needed. Because the initiation of a cell alters the
      # workbook, we do it on a clone.
      wanted_dims <- unname(unlist(dims))
      need_dims   <- wanted_dims[!wanted_dims %in% self$worksheets[[sheet]]$sheet_data$cc$r]
      if (length(need_dims)) # could be enough to pass wanted_dims
        temp <- self$clone()$.__enclos_env__$private$do_cell_init(sheet, dims)
      else
        temp <- self

      sel <- temp$worksheets[[sheet]]$sheet_data$cc$r %in% wanted_dims

      sty <- temp$worksheets[[sheet]]$sheet_data$cc[sel, c("r", "c_s")]

      x <- sty$c_s
      names(x) <- sty$r
      x
    },

    #' @description set sheet style
    #' @param style style
    #' @return The `wbWorksheetObject`, invisibly
    set_cell_style = function(sheet = current_sheet(), dims, style) {

      if (length(dims) == 1 && grepl(":|;|,", dims))
        dims <- dims_to_dataframe(dims, fill = TRUE)
      sheet <- private$get_sheet_index(sheet)

      if (all(grepl("[A-Za-z]", style))) {
        if (is_dims(style)) {
          styid <- self$get_cell_style(dims = style, sheet = sheet)
        } else {
          styid <- self$styles_mgr$get_xf_id(style)
        }
      } else {
        styid <- style
      }

      private$do_cell_init(sheet, dims)

      # if a range is passed (e.g. "A1:B2") we need to get every cell
      dims <- unname(unlist(dims))

      sel <- self$worksheets[[sheet]]$sheet_data$cc$r %in% dims

      self$worksheets[[sheet]]$sheet_data$cc$c_s[sel] <- styid

      invisible(self)
    },

    #' @description set style across columns and/or rows
    #' @param sheet sheet
    #' @param style style
    #' @param cols cols
    #' @param rows rows
    #' @return The `wbWorkbook` object
    set_cell_style_across = function(sheet = current_sheet(), style, cols = NULL, rows = NULL) {

      sheet <- private$get_sheet_index(sheet)
      if (all(grepl("[A-Za-z]", style))) {
        if (is_dims(style)) {
          styid <- self$get_cell_style(dims = style, sheet = sheet)
        } else {
          styid <- self$styles_mgr$get_xf_id(style)
        }
      } else {
        styid <- style
      }

      if (!is.null(rows)) {
        if (is.character(rows)) # row2int
          rows <- as.integer(dims_to_rowcol(rows)[["row"]])

        dims  <- wb_dims(rows, "A")
        cells <- unname(unlist(dims_to_dataframe(dims, fill = TRUE)))
        cc    <- self$worksheets[[sheet]]$sheet_data$cc

        cells <- cells[!cells %in% cc$r]
        if (length(cells) > 0) {
          private$do_cell_init(sheet, dims)
          self$set_cell_style(sheet = sheet, dims = cells, style = styid)
        }

        rows_df <- self$worksheets[[sheet]]$sheet_data$row_attr
        sel     <- rows_df$r %in% as.character(as.integer(rows))

        rows_df$customFormat[sel] <- "1"
        rows_df$s[sel]            <- styid
        self$worksheets[[sheet]]$sheet_data$row_attr <- rows_df

      }

      if (!is.null(cols)) {

        cols <- col2int(cols)

        cols_df <- wb_create_columns(self, sheet, cols)
        sel <- cols_df$min %in% as.character(cols)
        cols_df$style[sel] <- styid
        self$worksheets[[sheet]]$fold_cols(cols_df)

      }

      invisible(self)
    },

    #' @description set sheet style
    #' @param name name
    #' @param font_name,font_size optional else the default of the theme
    #' @return The `wbWorkbook`, invisibly
    add_named_style = function(
      sheet = current_sheet(),
      dims = "A1",
      name = "Normal",
      font_name = NULL,
      font_size = NULL
    ) {

      if (is.null(font_name)) font_name <- self$get_base_font()$name$val
      if (is.null(font_size)) font_size <- self$get_base_font()$size$val

      # if required initialize the cell style
      self$styles_mgr$init_named_style(name, font_name, font_size)

      ids <- self$styles_mgr$getstyle_ids(name)

      border_id <- ids[["borderId"]]
      fill_id   <- ids[["fillId"]]
      font_id   <- ids[["fontId"]]
      numfmt_id <- ids[["numFmtId"]]
      title_id  <- ids[["titleId"]]

      self$add_cell_style(
        dims       = dims,
        border_id  = border_id,
        fill_id    = fill_id,
        font_id    = font_id,
        num_fmt_id = numfmt_id,
        xf_id      = title_id
      )
      invisible(self)
    },

    #' @description create dxfs style
    #' These styles are used with conditional formatting and custom table styles
    #' @param name the style name
    #' @param font_name the font name
    #' @param font_size the font size
    #' @param font_color the font color (a `wb_color()` object)
    #' @param num_fmt the number format
    #' @param border logical if borders are applied
    #' @param border_color the border color
    #' @param border_style the border style
    #' @param bg_fill any background fill
    #' @param gradient_fill any gradient fill
    #' @param text_bold logical if text is bold
    #' @param text_italic logical if text is italic
    #' @param text_underline logical if text is underlined
    #' @param ... additional arguments passed to `create_dxfs_style()`
    #' @return The `wbWorksheetObject`, invisibly
    #' @export
    add_dxfs_style = function(
      name,
      font_name      = NULL,
      font_size      = NULL,
      font_color     = NULL,
      num_fmt        = NULL,
      border         = NULL,
      border_color   = wb_color(getOption("openxlsx2.borderColor", "black")),
      border_style   = getOption("openxlsx2.borderStyle", "thin"),
      bg_fill        = NULL,
      gradient_fill  = NULL,
      text_bold      = NULL,
      text_italic    = NULL,
      text_underline = NULL,
      ...
    ) {

      standardize(...)

      xml_style <- create_dxfs_style(
        font_name      = font_name,
        font_size      = font_size,
        font_color     = font_color,
        num_fmt        = num_fmt,
        border         = border,
        border_color   = border_color,
        border_style   = border_style,
        bg_fill        = bg_fill,
        gradient_fill  = gradient_fill,
        text_bold      = text_bold,
        text_italic    = text_italic,
        text_underline = text_underline,
        ...            = ...
      )

      got <- self$styles_mgr$get_dxf_id(name)

      if (!is.null(got) && !is.na(got))
        warning("dxfs style names should be unique")

      self$add_style(xml_style, name)

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
    #' @param sparklines sparkline created by `create_sparkline()`
    add_sparklines = function(
      sheet = current_sheet(),
      sparklines
    ) {
      sheet <- private$get_sheet_index(sheet)
      sparklines <- replace_waiver(sparklines, wb = self)
      self$worksheets[[sheet]]$add_sparklines(sparklines)
      invisible(self)
    },

    #' @description Ignore error on worksheet
    #' @param calculated_column calculatedColumn
    #' @param empty_cell_reference emptyCellReference
    #' @param eval_error evalError
    #' @param formula formula
    #' @param formula_range formulaRange
    #' @param list_data_validation listDataValidation
    #' @param number_stored_as_text numberStoredAsText
    #' @param two_digit_text_year twoDigitTextYear
    #' @param unlocked_formula unlockedFormula
    add_ignore_error = function(
      sheet                 = current_sheet(),
      dims                  = "A1",
      calculated_column     = FALSE,
      empty_cell_reference  = FALSE,
      eval_error            = FALSE,
      formula               = FALSE,
      formula_range         = FALSE,
      list_data_validation  = FALSE,
      number_stored_as_text = FALSE,
      two_digit_text_year   = FALSE,
      unlocked_formula      = FALSE,
      ...
    ) {
      standardize_case_names(...)
      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$ignore_error(
        dims               = dims,
        calculatedColumn   = calculated_column,
        emptyCellReference = empty_cell_reference,
        evalError          = eval_error,
        formula            = formula,
        formulaRange       = formula_range,
        listDataValidation = list_data_validation,
        numberStoredAsText = number_stored_as_text,
        twoDigitTextYear   = two_digit_text_year,
        unlockedFormula    = unlocked_formula
      )
      invisible(self)
    },

    #' @description add sheetview
    #' @param color_id,default_grid_color Integer: A color, default is 64
    #' @param right_to_left Logical: if TRUE column ordering is right  to left
    #' @param show_formulas Logical: if TRUE cell formulas are shown
    #' @param show_grid_lines Logical: if TRUE the worksheet grid is shown
    #' @param show_outline_symbols Logical: if TRUE outline symbols are shown
    #' @param show_row_col_headers Logical: if TRUE row and column headers are shown
    #' @param show_ruler Logical: if TRUE a ruler is shown in page layout view
    #' @param show_white_space Logical: if TRUE margins are shown in page layout view
    #' @param show_zeros Logical: if FALSE cells containing zero are shown blank if !showFormulas
    #' @param tab_selected Integer: zero vector indicating the selected tab
    #' @param top_left_cell Cell: the cell shown in the top left corner / or top right with rightToLeft
    #' @param view View: "normal", "pageBreakPreview" or "pageLayout"
    #' @param window_protection Logical: if TRUE the panes are protected
    #' @param workbook_view_id integer: Pointing to some other view inside the workbook
    #' @param zoom_scale,zoom_scale_normal,zoom_scale_page_layout_view,zoom_scale_sheet_layout_view Integer: the zoom scale should be between 10 and 400. These are values for current, normal etc.
    #' @return The `wbWorksheetObject`, invisibly
    set_sheetview = function(
      sheet                    = current_sheet(),
      color_id                     = NULL,
      default_grid_color           = NULL,
      right_to_left                = NULL,
      show_formulas                = NULL,
      show_grid_lines              = NULL,
      show_outline_symbols         = NULL,
      show_row_col_headers         = NULL,
      show_ruler                   = NULL,
      show_white_space             = NULL,
      show_zeros                   = NULL,
      tab_selected                 = NULL,
      top_left_cell                = NULL,
      view                         = NULL,
      window_protection            = NULL,
      workbook_view_id             = NULL,
      zoom_scale                   = NULL,
      zoom_scale_normal            = NULL,
      zoom_scale_page_layout_view  = NULL,
      zoom_scale_sheet_layout_view = NULL,
      ...
    ) {
      sheet <- private$get_sheet_index(sheet)
      self$worksheets[[sheet]]$set_sheetview(
        color_id                     = color_id,
        default_grid_color           = default_grid_color,
        right_to_left                = right_to_left,
        show_formulas                = show_formulas,
        show_grid_lines              = show_grid_lines,
        show_outline_symbols         = show_outline_symbols,
        show_row_col_headers         = show_row_col_headers,
        show_ruler                   = show_ruler,
        show_white_space             = show_white_space,
        show_zeros                   = show_zeros,
        tab_selected                 = tab_selected,
        top_left_cell                = top_left_cell,
        view                         = view,
        window_protection            = window_protection,
        workbook_view_id             = workbook_view_id,
        zoom_scale                   = zoom_scale,
        zoom_scale_normal            = zoom_scale_normal,
        zoom_scale_page_layout_view  = zoom_scale_page_layout_view,
        zoom_scale_sheet_layout_view = zoom_scale_sheet_layout_view,
        ...                          = ...
      )
      invisible(self)
    },

    #' @description add person to workbook
    #' @param name name
    #' @param id id
    #' @param user_id user_id
    #' @param provider_id provider_id
    add_person = function(
      name        = NULL,
      id          = NULL,
      user_id     = NULL,
      provider_id = "None"
    ) {

      if (is.null(name))    name    <- Sys.getenv("USERNAME", Sys.getenv("USER"))
      if (is.null(id))      id      <- st_guid()
      if (is.null(user_id)) user_id <- st_userid()

      xml_person <- xml_node_create(
        "person",
        xml_attributes = c(
          displayName = name,
          id          = id,
          userId      = user_id,
          providerId  = "None"
        )
      )

      options("openxlsx2.thread_id" = id)

      if (is.null(self$persons)) {
        self$persons <- xml_node_create(
          "personList",
          xml_attributes = c(
            `xmlns`   = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
            `xmlns:x` = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          )
        )

        self$append(
          "workbook.xml.rels",
          "<Relationship Id=\"rId5\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/person\" Target=\"persons/person.xml\"/>"
        )

        self$append(
          "Content_Types",
          "<Override PartName=\"/xl/persons/person.xml\" ContentType=\"application/vnd.ms-excel.person+xml\"/>"
        )
      }

      self$persons <- xml_add_child(self$persons, xml_person)

      invisible(self)
    },

    #' @description description get person
    #' @param name name
    get_person = function(name = NULL) {
      persons <- rbindlist(xml_attr(self$persons, "personList", "person"))
      if (!is.null(name)) persons <- persons[persons$displayName == name, ]
      persons
    },

    #' @description description get active sheet
    get_active_sheet = function() {
      at <- rbindlist(xml_attr(self$workbook$bookViews, "bookViews", "workbookView"))$activeTab
      # return c index as R index
      as.numeric(at) + 1
    },

    #' @description description set active sheet
    set_active_sheet = function(sheet = current_sheet()) {
      sheet <- private$get_sheet_index(sheet)
      self$set_bookview(active_tab = sheet - 1L)
    },

    #' @description description get selected sheets
    get_selected = function() {
      len <- length(self$sheet_names)
      sv <- vector("list", length = len)

      for (i in seq_len(len)) {
        sv[[i]] <- xml_node(self$worksheets[[i]]$sheetViews, "sheetViews", "sheetView")
      }

      # print(sv)
      z <- rbindlist(xml_attr(sv, "sheetView"))
      z$names <- self$get_sheet_names(escape = TRUE)
      z
    },

    #' @description set selected sheet
    set_selected = function(sheet = current_sheet()) {

      sheet <- private$get_sheet_index(sheet)

      for (i in seq_along(self$sheet_names)) {
        xml_attr <- i == sheet
        self$worksheets[[i]]$set_sheetview(tab_selected = xml_attr)
      }

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
        warning("Fixing: removing illegal characters found in sheet name. See ?openxlsx2::clean_worksheet_name.")
        sheet <- replace_illegal_chars(sheet)
      }

      if (!nzchar(sheet)) {
        warning("Fixing: sheet name must contain at least 1 character.")
        sheet <- paste("Sheet", length(self$sheet_names) + 1)
      }

      if (nchar(sheet) > 31) {
        warning("Fixing: shortening sheet name to 31 characters.")
        sheet <- stringi::stri_sub(sheet, 1, 31)
        if (anyDuplicated(c(sheet, self$sheet_names)))
          stop(
            "Cannot shorten sheet name to a unique string. ",
            "Please provide a unique sheetname with maximum 31 characters."
          )
      }

      if (tolower(sheet) %in% self$sheet_names) {
        warning('Attempted to add a worksheet that is invalid or already exists.\n',
                'Fixing: a sheet with name "', sheet, '" already exists. Creating a unique sheetname"', call. = FALSE)
        ## We simply append (1), while spreadsheet software would increase
        ## the integer as: Sheet, Sheet (1), Sheet (2) etc.
        sheet <- paste(sheet, "(1)")
      }

      assign("sheet", sheet, parent.frame())
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

    add_media = function(
      file
    ) {

      imageType <- file_ext2(file)
      mediaNo   <- length(self$media) + 1L

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

      ## write file path to media slot to copy across on save
      tmp <- file
      names(tmp) <- stringi::stri_join("image", mediaNo, ".", imageType)
      self$append("media", tmp)

      invisible(self)
    },

    get_drawingsref = function() {
      has_drawing <- grep("drawings", self$worksheets_rels)

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

      # beg vml loop
      for (i in seq_along(self$vml)) {

        ## write vml output
        if (self$vml[[i]] != "") {
          write_file(
              head = "",
              body = pxml(self$vml[[i]]),
              tail = "",
              fl = file.path(dir, sprintf("vmlDrawing%s.vml", i))
          )

          if (!is.null(unlist(self$vml_rels)) && length(self$vml_rels) >= i && !all(self$vml_rels[[i]] == "")) {
            write_file(
              head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
              body = pxml(self$vml_rels[[i]]),
              tail = '</Relationships>',
              fl = file.path(dir_rel, sprintf("vmlDrawing%s.vml.rels", i))
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
      xlworksheetsRelsDir,
      use_pugixml_export
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
              fl = file.path(xlchartsDir, stringi::stri_join("chart", crt, ".xml"))
            )
          }

          if (self$charts$chartEx[crt] != "") {
            ct <- c(ct, sprintf('<Override PartName="/xl/charts/chartEx%s.xml" ContentType="application/vnd.ms-office.chartex+xml"/>', crt))

            write_file(
              body = self$charts$chartEx[crt],
              fl = file.path(xlchartsDir, stringi::stri_join("chartEx", crt, ".xml"))
            )
          }

          if (self$charts$colors[crt] != "") {
            ct <- c(ct, sprintf('<Override PartName="/xl/charts/colors%s.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>', crt))

            write_file(
              body = self$charts$colors[crt],
              fl = file.path(xlchartsDir, stringi::stri_join("colors", crt, ".xml"))
            )
          }

          if (self$charts$style[crt] != "") {
            ct <- c(ct, sprintf('<Override PartName="/xl/charts/style%s.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>', crt))

            write_file(
              body = self$charts$style[crt],
              fl = file.path(xlchartsDir, stringi::stri_join("style", crt, ".xml"))
            )
          }

          if (self$charts$rels[crt] != "") {
            write_file(
              body = self$charts$rels[crt],
              fl = file.path(xlchartsRelsDir, stringi::stri_join("chart", crt, ".xml.rels"))
            )
          }

          if (self$charts$relsEx[crt] != "") {
            write_file(
              body = self$charts$relsEx[crt],
              fl = file.path(xlchartsRelsDir, stringi::stri_join("chartEx", crt, ".xml.rels"))
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
            head = "",
            body = pxml(self$drawings[[i]]),
            tail = "",
            fl = file.path(xldrawingsDir, stringi::stri_join("drawing", i, ".xml"))
          )
          if (!all(self$drawings_rels[[i]] == "")) {
            write_file(
              head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
              body = pxml(self$drawings_rels[[i]]),
              tail = '</Relationships>',
              fl = file.path(xldrawingsRelsDir, stringi::stri_join("drawing", i, ".xml.rels"))
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
            fl = file.path(chartSheetDir, stringi::stri_join("sheet", i, ".xml"))
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
          # ws <- self$worksheets[[i]]
          hasHL <- length(self$worksheets[[i]]$hyperlinks) > 0

          prior <- self$worksheets[[i]]$get_prior_sheet_data()
          post <-  self$worksheets[[i]]$get_post_sheet_data()

          if (use_pugixml_export) {
            # failsaves. check that all rows and cells
            # are available and in the correct order
            if (!is.null(self$worksheets[[i]]$sheet_data$cc)) {

              self$worksheets[[i]]$sheet_data$cc$r <- with(
                self$worksheets[[i]]$sheet_data$cc,
                stringi::stri_join(c_r, row_r)
              )
              cc <- self$worksheets[[i]]$sheet_data$cc
              # prepare data for output

              # there can be files, where row_attr is incomplete because a row
              # is lacking any attributes (presumably was added before saving)
              # still row_attr is what we want!

              rows_attr <- self$worksheets[[i]]$sheet_data$row_attr
              self$worksheets[[i]]$sheet_data$row_attr <- rows_attr[order(as.numeric(rows_attr[, "r"])), ]

              cc_rows <- self$worksheets[[i]]$sheet_data$row_attr$r
              # c("row_r", "c_r",  "r", "v", "c_t", "c_s", "c_cm", "c_ph", "c_vm", "f", "f_attr", "is")
              cc <- cc[cc$row_r %in% cc_rows, ]

              sort_key <- as.numeric(cc$row_r) * 16384L + col2int(cc$c_r)
              self$worksheets[[i]]$sheet_data$cc <- cc[order(sort_key), ]
              rm(cc)
            } else {
              self$worksheets[[i]]$sheet_data$row_attr <- NULL
              self$worksheets[[i]]$sheet_data$cc <- NULL
            }
          }

          ws_file <- file.path(xlworksheetsDir, sprintf("sheet%s.xml", i))

          if (use_pugixml_export) {

            # create entire sheet prior to writing it
            sheet_xml <- write_worksheet(
              prior      = prior,
              post       = post,
              sheet_data = self$worksheets[[i]]$sheet_data
            )
            write_xmlPtr(doc = sheet_xml, fl = ws_file)

          } else {

            if (grepl("</worksheet>", prior))
              prior <- substr(prior, 1, nchar(prior) - 13) # remove " </worksheet>"

            write_worksheet_slim(
              sheet_data = self$worksheets[[i]]$sheet_data,
              prior      = prior,
              post       = post,
              fl         = ws_file
            )

          }

          ## write worksheet rels
          if (length(self$worksheets_rels[[i]])) {
            ws_rels <- self$worksheets_rels[[i]]

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

      ct
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

      if (!grepl("^[\\p{L}_][^\\s]*$", name, perl = TRUE))
        stop("named region must begin with a letter or an underscore and not contain whitespace(s).")

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
        comment           = comment,
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
      sqref <- stringi::stri_join(
        get_cell_refs(data.frame(x = c(startRow, endRow), y = c(startCol, endCol), stringsAsFactors = FALSE)),
        collapse = ":"
      )

      nms <- c(names(self$worksheets[[sheet]]$conditionalFormatting), sqref)
      dxfId <- max(dxfId, 0L)

      priority <- max(0,
        as.integer(
          openxlsx2:::rbindlist(
            xml_attr(self$worksheets[[sheet]]$conditionalFormatting, "cfRule")
          )$priority
        )
      ) + 1L

      # big switch statement
      cfRule <- switch(
        type,

        ## colorScale ----
        colorScale = cf_create_colorscale(priority, formula, values),

        ## dataBar ----
        dataBar = cf_create_databar(priority, self$worksheets[[sheet]]$extLst, formula, params, sqref, values),

        ## expression ----
        expression = cf_create_expression(priority, dxfId, formula),

        ## duplicatedValues ----
        duplicatedValues = cf_create_duplicated_values(priority, dxfId),

        ## containsText ----
        containsText = cf_create_contains_text(priority, dxfId, sqref, values),

        ## notContainsText ----
        notContainsText = cf_create_not_contains_text(priority, dxfId, sqref, values),

        ## beginsWith ----
        beginsWith = cf_begins_with(priority, dxfId, sqref, values),

        ## endsWith ----
        endsWith = cf_ends_with(priority, dxfId, sqref, values),

        ## between ----
        between = cf_between(priority, dxfId, formula),

        ## topN ----
        topN = cf_top_n(priority, dxfId, values),

        ## bottomN ----
        bottomN = cf_bottom_n(priority, dxfId, values),

        ## uniqueValues ---
        uniqueValues = cf_unique_values(priority, dxfId),

        ## iconSet ----
        iconSet = cf_icon_set(priority, self$worksheets[[sheet]]$extLst, sqref, values, params),

        ## containsErrors ----
        containsErrors = cf_iserror(priority, dxfId, sqref),

        ## notContainsErrors ----
        notContainsErrors = cf_isnoerror(priority, dxfId, sqref),

        ## containsBlanks ----
        containsBlanks = cf_isblank(priority, dxfId, sqref),

        ## notContainsBlanks ----
        notContainsBlanks = cf_isnoblank(priority, dxfId, sqref),

        # do we have a match.arg() anywhere or will it just be showned in this switch()?
        stop("type `", type, "` is not a valid formatting rule")
      )

      # dataBar needs additional extLst
      if (!is.null(attr(cfRule, "extLst"))) {
        # self$worksheets[[sheet]]$extLst <- read_xml(attr(cfRule, "extLst"), pointer = FALSE)
        self$worksheets[[sheet]]$.__enclos_env__$private$do_append_x14(attr(cfRule, "extLst"), "x14:conditionalFormatting", "x14:conditionalFormattings")
      }

      if (length(cfRule)) {
        private$append_sheet_field(sheet, "conditionalFormatting", read_xml(cfRule, pointer = FALSE))
        names(self$worksheets[[sheet]]$conditionalFormatting) <- nms
      }
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
      richDataInd      <- grep("richData",                                   self$workbook.xml.rels)


      ## Reordering of workbook.xml.rels
      ## don't want to re-assign rIds for pivot tables or slicer caches
      pivotNode        <- grep("pivotCache/pivotCacheDefinition[0-9]+.xml", self$workbook.xml.rels, value = TRUE)
      slicerNode       <- grep("slicerCache[0-9]+.xml",                     self$workbook.xml.rels, value = TRUE)
      timelineNode     <- grep("timelineCache[0-9]+.xml",                   self$workbook.xml.rels, value = TRUE)

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
          calcChainInd,
          richDataInd
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

      self$append("workbook.xml.rels", c(pivotNode, slicerNode, timelineNode))

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

      if (!is.null(self$featurePropertyBag)) {
        self$append("workbook.xml.rels",
          sprintf(
            '<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag" Target="featurePropertyBag/featurePropertyBag.xml"/>',
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


      ## re-assign tab_selected
      state <- rep.int("visible", nSheets)
      hidden <- grepl("hidden", self$workbook$sheets)
      state[hidden] <- "hidden"
      visible_sheet_index <- which(!hidden)[1] # first visible

      if (is.null(self$workbook$bookViews))
        self$set_bookview(
          x_window      = 0,
          y_window      = 0,
          window_width  = 13125,
          window_height = 13125,
          first_sheet   = visible_sheet_index - 1L,
          active_tab    = visible_sheet_index - 1L
        )

      # Failsafe: hidden sheet can not be selected.
      if (any(hidden)) {
        for (i in which(hidden)) {
          self$worksheets[[i]]$set_sheetview(tab_selected = FALSE)
        }
      }

      ## update workbook r:id to match reordered workbook.xml.rels externalLink element
      if (length(extRefInds)) {
        newInds <- seq_along(extRefInds) + length(sheetInds)
        self$workbook$externalReferences <- stringi::stri_join(
          "<externalReferences>",
          stringi::stri_join(sprintf('<externalReference r:id=\"rId%s\"/>', newInds), collapse = ""),
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

      if (is.null(self$worksheets[[sheet]]$sheet_data$cc)) {
        # everythings missing, we can safely write data

        self$add_data(
          sheet = sheet,
          x = dims_to_dataframe(dims),
          na.strings = NULL,
          col_names = FALSE,
          dims = dims
        )

      } else {
        # there are some cells already available, we have to create the missing cells

        need_cells <- dims
        if (length(need_cells) == 1 && grepl(":|;|,", need_cells))
          need_cells <- dims_to_dataframe(dims, fill = TRUE)

        exp_cells <- unname(unlist(need_cells[need_cells != ""]))
        got_cells <- self$worksheets[[sheet]]$sheet_data$cc$r

        # initialize cell
        if (!all(exp_cells %in% got_cells)) {
            missing_cells <- exp_cells[!exp_cells %in% got_cells]
            self <- initialize_cell(self, sheet = sheet, new_cells = missing_cells)
        }

      }

      invisible(self)
    }
  )
)
