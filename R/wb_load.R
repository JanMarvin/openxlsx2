#' Load an existing .xlsx, .xlsm or .xlsb file
#'
#' `wb_load()` returns a [wbWorkbook] object conserving the content of the
#' original input file, including data, styles, media. This workbook can be
#' modified, read from, and be written back into a xlsx file.
#'
#' @details
#' If a specific `sheet` is selected, the workbook will still contain sheets
#' for all worksheets. The argument `sheet` and `data_only` are used internally
#' by [wb_to_df()] to read from a file with minimal changes. They are not
#' specifically designed to create rudimentary but otherwise fully functional
#' workbooks. It is possible to import with
#' `wb_load(data_only = TRUE, sheet = NULL)`. In this way, only a workbook
#' framework is loaded without worksheets or data. This can be useful if only
#' some workbook properties are of interest.
#'
#' There are some internal arguments that can be passed to wb_load, which are
#' used for development. The `debug` argument allows debugging of `xlsb` files
#' in particular. With `calc_chain` it is possible to maintain the calculation
#' chain. The calculation chain is used by spreadsheet software to determine
#' the order in which formulas are evaluated. Removing the calculation chain
#' has no known effect. The calculation chain is created the next time the
#' worksheet is loaded into the spreadsheet. Keeping the calculation chain
#' could only shorten the loading time in said software. Unfortunately, if a
#' cell is added to the worksheet, the calculation chain may block the
#' worksheet as the formulas will not be evaluated again until each individual
#' cell with a formula is selected in the spreadsheet software and the Enter
#' key is pressed manually. It is therefore strongly recommended not to
#' activate this function.
#'
#' In rare cases, a warning is issued when loading an xlsx file that an xml
#' namespace has been removed from xml files. This refers to the internal
#' structure of the loaded xlsx file. Certain xlsx files created by third-party
#' applications contain a namespace (usually x). This namespace is not required
#' for the file to work in spreadsheet software and is not expected by
#' `openxlsx2`. It is therefore removed when the file is loaded into a
#' workbook. Removal is generally considered safe, but the feature is still not
#' commonly observed, hence the warning.
#'
#' Initial support for binary openxml files (`xlsb`) has been added to the
#' package. We parse the binary file format into pseudo-openxml files that we
#' can import. Therefore, once imported, it is possible to interact with the
#' file as if it had been provided in xlsx file format in the first place. This
#' parsing into pseudo xml files is of course slower than reading directly from
#' the binary file. Our implementation is also still missing some functions:
#' some array formulas are not yet correct, conditional formatting and data
#' validation are not implemented, nor are pivot tables and slicers.
#'
#' @param file A path to an existing .xlsx, .xlsm or .xlsb file
#' @param sheet optional sheet parameter. if this is applied, only the selected
#'   sheet will be loaded. This can be a numeric, a string or `NULL`.
#' @param data_only mode to import if only a data frame should be returned. This
#'   strips the `wbWorkbook` to a bare minimum.
#' @param ... additional arguments
#' @return A Workbook object.
#' @examples
#' ## load existing workbook
#' fl <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
#' wb <- wb_load(file = fl)
#' @export
wb_load <- function(
    file,
    sheet,
    data_only = FALSE,
    ...
) {

  calc_chain <- list(...)$calc_chain
  debug      <- list(...)$debug
  xlsx_file  <- list(...)$xlsx_file
  standardize_case_names(...)

  if (is.null(calc_chain)) calc_chain <- FALSE
  if (is.null(debug)) debug <- FALSE

  if (!is.null(xlsx_file)) {
    .Deprecated(old = "xlsx_file", new = "file", package = "openxlsx2")
    file <- xlsx_file %||% file
  }

  file <- getFile(file)

  if (!file.exists(file)) {
    stop("File does not exist.")
  }

  ## create temp dir
  xmlDir <- temp_dir("_openxlsx_wb_load")

  # do not unlink after loading
  # on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)

  ## Unzip files to temp directory
  xmlFiles <- withCallingHandlers(
    utils::unzip(file, exdir = xmlDir),
    warning = function(w) {
      msg <- paste("Unable to open and load file: ", file)
      stop(msg, call. = FALSE)
    }
  )

  # we need to read the files in human order: 1, 2, 10 and not 1, 10, 2.
  ordr <- stringi::stri_order(xmlFiles, opts_collator = stringi::stri_opts_collator(numeric = TRUE))
  xmlFiles <- xmlFiles[ordr]

  wb <- wb_workbook()
  wb$path <- file

  # There is one known file in #1194. this file has lower case folders, while
  # the references in the file are the usual camel case.
  needs_lower <- ifelse(any(grepl("\\[content_types\\].xml$", xmlFiles)), TRUE, FALSE)

  grep_xml <- function(pattern, perl = TRUE, value = TRUE, to_lower = needs_lower, ...) {
    # targets xmlFiles; has presents
    if (to_lower) pattern <- tolower(pattern)
    grep(pattern, xmlFiles, perl = perl, value = value, ...)
  }

  ## We have found a zip file, but it must not necessarily be a spreadsheet
  ContentTypesXML   <- grep_xml("\\[Content_Types\\].xml$")
  worksheetsXML     <- grep_xml("/worksheets/sheet[0-9]+")

  if ((length(ContentTypesXML) == 0 || length(worksheetsXML) == 0) && !debug) {
    msg <- paste("File does not appear to be xlsx, xlsm or xlsb: ", file)
    stop(msg)
  }

  # relsXML           <- grep_xml("_rels/.rels$")

  appXML            <- grep_xml("app.xml$")
  coreXML           <- grep_xml("core.xml$")
  customXML         <- grep_xml("custom.xml$")

  customXmlDir      <- grep_xml("customXml/")
  docMetadataXML    <- grep_xml("docMetadata/")

  workbookBIN       <- grep_xml("workbook.bin$")
  workbookXML       <- grep_xml("workbook.xml$")
  workbookBINRels   <- grep_xml("workbook.bin.rels")
  workbookXMLRels   <- grep_xml("workbook.xml.rels")

  drawingsXML       <- grep_xml("drawings/drawing[0-9]+.xml$")

  stylesBIN         <- grep_xml("styles.bin$")
  stylesXML         <- grep_xml("styles.xml$")
  sharedStringsBIN  <- grep_xml("sharedStrings.bin$")
  sharedStringsXML  <- grep_xml("sharedStrings.xml$")
  metadataXML       <- grep_xml("metadata.xml$")
  themeXML          <- grep_xml("theme[0-9]+.xml$")
  drawingRelsXML    <- grep_xml("drawing[0-9]+.xml.rels$")
  sheetRelsBIN      <- grep_xml("sheet[0-9]+.bin.rels$")
  sheetRelsXML      <- grep_xml("sheet[0-9]+.xml.rels$")
  media             <- grep_xml("image[0-9]+.[a-z]+$")
  vmlDrawingXML     <- grep_xml("drawings/vmlDrawing[0-9]+\\.vml$")
  vmlDrawingRelsXML <- grep_xml("vmlDrawing[0-9]+.vml.rels$")
  calcChainXML      <- grep_xml("xl/calcChain.xml")
  embeddings        <- grep_xml("xl/embeddings")
  activeX           <- grep_xml("xl/activeX")

  # comments
  commentsBIN       <- grep_xml("xl/comments[0-9]+\\.bin")
  commentsXML       <- grep_xml("xl/comments[0-9]+\\.xml")
  personXML         <- grep_xml("xl/persons/person.xml$")
  threadCommentsXML <- grep_xml("xl/threadedComments/threadedComment[0-9]+\\.xml")

  # charts
  chartsXML         <- grep_xml("xl/charts/chart[0-9]+\\.xml$")
  chartExsXML       <- grep_xml("xl/charts/chartEx[0-9]+\\.xml$")
  chartsXML_colors  <- grep_xml("xl/charts/colors[0-9]+\\.xml$")
  chartsXML_styles  <- grep_xml("xl/charts/style[0-9]+\\.xml$")
  chartsRels        <- grep_xml("xl/charts/_rels/chart[0-9]+.xml.rels")
  chartExsRels      <- grep_xml("xl/charts/_rels/chartEx[0-9]+.xml.rels")
  chartSheetsXML    <- grep_xml("xl/chartsheets/sheet[0-9]+")

  # tables
  tablesBIN         <- grep_xml("tables/table[0-9]+.bin$")
  tablesXML         <- grep_xml("tables/table[0-9]+.xml$")
  tableRelsXML      <- grep_xml("table[0-9]+.xml.rels$")
  queryTablesXML    <- grep_xml("queryTable[0-9]+.xml$")

  # connections
  connectionsXML    <- grep_xml("connections.xml$")
  extLinksBIN       <- grep_xml("externalLink[0-9]+.bin$")
  extLinksXML       <- grep_xml("externalLink[0-9]+.xml$")
  extLinksRelsBIN   <- grep_xml("externalLink[0-9]+.bin.rels$")
  extLinksRelsXML   <- grep_xml("externalLink[0-9]+.xml.rels$")

  # form control
  ctrlPropsXML      <- grep_xml("ctrlProps/ctrlProp[0-9]+.xml")

  # pivot tables
  pivotTableXML     <- grep_xml("pivotTable[0-9]+.xml$")
  pivotTableRelsXML <- grep_xml("pivotTable[0-9]+.xml.rels$")
  pivotDefXML       <- grep_xml("pivotCacheDefinition[0-9]+.xml$")
  pivotDefRelsXML   <- grep_xml("pivotCacheDefinition[0-9]+.xml.rels$")
  pivotCacheRecords <- grep_xml("pivotCacheRecords[0-9]+.xml$")

  # rich data
  rdrichvalue       <- grep_xml("richData/rdrichvalue.xml")
  rdrichvaluestr    <- grep_xml("richData/rdrichvaluestructure.xml")
  rdRichValueTypes  <- grep_xml("richData/rdRichValueTypes.xml")
  richValueRel      <- grep_xml("richData/richValueRel.xml")
  richValueRelrels  <- grep_xml("richData/_rels/richValueRel.xml.rels")

  ## slicers
  slicerXML         <- grep_xml("slicer[0-9]+.xml$")
  slicerCachesXML   <- grep_xml("slicerCache[0-9]+.xml$")

  ## timelines
  timelineXML         <- grep_xml("timeline[0-9]+.xml$")
  timelineCachesXML   <- grep_xml("timelineCache[0-9]+.xml$")

  ## VBA Macro
  vbaProject        <- grep_xml("vbaProject\\.bin$")

  ## feature property bag
  featureProperty   <- grep_xml("featurePropertyBag.xml$")

  cleanup_dir <- function(data_only) {
    grep_xml("media|vmlDrawing|customXml|embeddings|activeX|vbaProject", ignore.case = TRUE, invert = TRUE)
  }

  ## remove all EXCEPT media and charts
  if (!data_only) on.exit(
    unlink(
      # TODO: this removes all files, the folders remain
      cleanup_dir(data_only),
      recursive = TRUE, force = TRUE
    ),
    add = TRUE
  )

  file_folders <- unique(basename(dirname(xmlFiles)))
  known <- c(
    basename(xmlDir), "_rels", "activeX", "charts", "chartsheets",
    "ctrlProps", "customXml", "docMetadata", "docProps", "drawings",
    "embeddings", "externalLinks", "featurePropertyBag", "media",
    "persons", "pivotCache", "pivotTables", "printerSettings",
    "queryTables", "richData", "slicerCaches", "slicers", "tables",
    "theme", "threadedComments", "timelineCaches", "timelines",
    "worksheets", "xl", "[trash]"
  )
  unknown <- file_folders[!file_folders %in% known]
  # nocov start
  if (length(unknown)) {
    message <- paste0(
      "Found unknown folders in the input file:\n",
      paste(unknown, collapse = "\n"), "\n\n",
      "These folders are not yet processed by openxlsx2, but depending on what you want to do, this may not be fatal.\n",
      "Most likely a file with these folders has not yet been detected. If you want to contribute to the development of the package, maybe just to silence this warning, please open an issue about unknown content folders and, if possible, provide a file via our Github or by mail."
    )
    warning(message, call. = FALSE)
  }
  # nocov end

  # modifications for xlsb
  if (length(workbookBIN)) {

    if (file.info(file)$size > 100000) {
      message("importing larger workbook. please wait a moment")
    }

    workbookXML <- gsub(".bin$", ".xml", workbookBIN)
    if (debug) {
      print(workbookBIN)
      print(workbookXML)
    }
    workbook_bin(workbookBIN, workbookXML, debug)

    if (length(stylesBIN)) {
      stylesXML <- gsub(".bin$", ".xml", stylesBIN)
      styles_bin(stylesBIN, stylesXML, debug)
      # system(sprintf("cat %s", stylesXML))
      # system(sprintf("cp %s /tmp/styles.xml", stylesXML))
    }

    if (length(sharedStringsBIN)) {
      sharedStringsXML <- gsub(".bin$", ".xml", sharedStringsBIN)
      sharedstrings_bin(sharedStringsBIN, sharedStringsXML, debug)
      # system(sprintf("cat %s", sharedStringsXML))
      # system(sprintf("cp %s /tmp/sst.xml", sharedStringsXML))
    }

    if (!data_only && length(tablesBIN)) {
      tablesXML <- gsub(".bin$", ".xml", tablesBIN)
      for (i in seq_along(tablesXML))
        table_bin(tablesBIN[i], tablesXML[i], debug)
      # system(sprintf("cat %s", tablesXML))
      # system(sprintf("cp %s /tmp/tables.xml", tablesXML))
    }

    if (!data_only && length(chartSheetsXML)) {
      chartSheetsXML <- gsub(".bin$", ".xml", chartSheetsXML)
    }

    if (!data_only && length(commentsBIN)) {
      commentsXML <- gsub(".bin$", ".xml", commentsBIN)
      for (i in seq_along(commentsBIN)) {
        comments_bin(commentsBIN[i], commentsXML[i], debug)
        # system(sprintf("cat %s", commentsXML[i]))
        # system(sprintf("cp %s /tmp/tables.xml", tablesXML))
      }
    }

    if (!data_only && length(extLinksBIN)) {
      extLinksXML <- gsub(".bin$", ".xml", extLinksBIN)
      for (i in seq_along(extLinksBIN)) {
        externalreferences_bin(extLinksBIN[i], extLinksXML[i], debug)
        # system(sprintf("cat %s", commentsXML[i]))
        # system(sprintf("cp %s /tmp/tables.xml", tablesXML))
      }

      extLinksRelsXML <- extLinksRelsBIN
    }

    workbookXMLRels <- workbookBINRels
    sheetRelsXML    <- sheetRelsBIN
  }

  ## core
  if (!data_only && length(appXML)) {
    app_xml <- read_xml(appXML)
    nodes <- xml_node_name(app_xml, "Properties")
    app_list <- lapply(
      nodes,
      FUN = function(x) xml_node(app_xml, "Properties", x)
    )
    names(app_list) <- nodes
    wb$app[names(app_list)] <- app_list
  }

  if (!data_only && length(coreXML) == 1) {
    wb$core <- read_xml(coreXML, pointer = FALSE)
  }

  if (!data_only && length(customXML)) {
    wb$append("Content_Types", '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>')
    wb$custom <- read_xml(customXML, pointer = FALSE)
  }

  if (!data_only && length(docMetadataXML)) {

    # rels <- read_xml(relsXML)
    # rels_df <- rbindlist(xml_attr(rels, "Relationships", "Relationship"))

    if (any(basename2(docMetadataXML) != "LabelInfo.xml"))
      warning("unknown metadata file found")

    wb$docMetadata <- read_xml(docMetadataXML, pointer = FALSE)
  }

  nSheets <- length(worksheetsXML) + length(chartSheetsXML)

  ##
  if (!data_only && length(workbookXMLRels)) {
    workbookRelsXML <- xml_node(workbookXMLRels, "Relationships", "Relationship")

    # TODO we remove all of the relationships and add them back later on.
    # I am not sure why we do this. We should simply keep it the way it is,
    # and only modify it if we need to.

    # Currently we create a slim skeleton and add relationships if needed. An
    # alternative to this is shown in the snippets below. Though this leaves
    # the possibility that we still include folders we do not handle correctly.
    # Like ctrlProps #206.

    # Ideally we keep the entire relationship as loaded and check only for
    # unknown content. And workbook.xml.rels should be checked pre or post
    # writing the output.

    # workbookRelsXML <- relship_no(workbookRelsXML, "connections")
    # workbookRelsXML <- relship_no(workbookRelsXML, "externalLink")
    # workbookRelsXML <- relship_no(workbookRelsXML, "person")
    # # our pivotTable has rId 2000 + x, but it is updated in preSaveClean
    # workbookRelsXML <- relship_no(workbookRelsXML, "pivotCacheDefinition")
    # workbookRelsXML <- relship_no(workbookRelsXML, "slicerCache")
    # workbookRelsXML <- relship_no(workbookRelsXML, "chartsheet")
    # workbookRelsXML <- relship_no(workbookRelsXML, "worksheet")
    # workbookRelsXML <- relship_no(workbookRelsXML, "vbaProject")
    # workbookRelsXML <- relship_no(workbookRelsXML, "sheetMetadata")

    need_sharedStrings <- FALSE

    if (any(grepl("sharedStrings", workbookRelsXML)))
      need_sharedStrings <- TRUE

    wb$workbook.xml.rels <- genBaseWorkbook.xml.rels()

    if (need_sharedStrings) {
      wb$append(
        "workbook.xml.rels",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
      )
    }
  }

  ##
  if (!data_only && length(chartSheetsXML)) {
    workbookRelsXML <- grep("chartsheets/sheet", workbookRelsXML, fixed = TRUE, value = TRUE)

    sheetNo <- as.integer(regmatches(chartSheetsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml)", chartSheetsXML, perl = TRUE)))
    chartSheetsXML <- chartSheetsXML[order(sheetNo)]

    chartSheetsRelsXML <- grep_xml("xl/chartsheets/_rels")
    sheetNo2 <- as.integer(regmatches(chartSheetsRelsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml\\.rels)", chartSheetsRelsXML, perl = TRUE)))
    chartSheetsRelsXML <- chartSheetsRelsXML[order(sheetNo2)]
  }

  ## xl\
  ## xl\workbook
  if (length(workbookXML)) {

    # escape
    workbook_xml <- read_xml(workbookXML, escapes = TRUE)

    xml_name <- xml_node_name(workbook_xml)[1]

    if (xml_name != "workbook") {
      xml_ns <- stringi::stri_split_fixed(xml_name, ":")[[1]][[1]]
      op <- options("openxlsx2.namespace_xml" = xml_ns)
      on.exit(options(op), add = TRUE)
      msg <- paste0(
        "The `{%s}` namespace(s) has been removed from the xml files, for example:\n",
        "\t<%s:field> changed to:\n",
        "\t<field>\n",
        "See 'Details' in ?openxlsx2::wb_load() for more information."
      )
      warning(sprintf(msg, xml_ns, xml_ns))
      workbook_xml <- read_xml(workbookXML, escapes = TRUE)
    }

    wb$workbook$fileVersion <- xml_node(workbook_xml, "workbook", "fileVersion")
    wb$workbook$alternateContent <- xml_node(workbook_xml, "workbook", "mc:AlternateContent")
    wb$workbook$bookViews <- xml_node(workbook_xml, "workbook", "bookViews")

    sheets <- xml_attr(workbook_xml, "workbook", "sheets", "sheet")
    sheets <- rbindlist(sheets)

    # Usually the id variable is called `r:id`, but there is one known sheet
    # that has `d3p1:id`
    r_id <- names(sheets)[grepl(":id", names(sheets))]

    ## Some veryHidden sheets do not have a sheet content and their rId is empty.
    ## Such sheets need to be filtered out because otherwise their sheet names
    ## occur in the list of all sheet names, leading to a wrong association
    ## of sheet names with sheet indeces.
    sheets <- sheets[sheets[r_id] != "", ]

    # if wb_relsxml is not available, the workbook has no relationships, not
    # sure if this is possible
    if (length(workbookXMLRels)) {
      wb_rels_xml <- rbindlist(
        xml_attr(workbookXMLRels, "Relationships", "Relationship")
      )
    }

    sheets <- merge(
      sheets, wb_rels_xml,
      by.x = r_id, by.y = "Id",
      all.x = TRUE,
      all.y = FALSE,
      sort = FALSE
    )

    # if /xl/ is not in the path add it
    xl_path <- ifelse(grepl("/xl/", sheets$Target), "", "/xl/")

    ## sheetId does not mean sheet order. Sheet order is defined in the index position here
    ## sheet rId links to the workbook.xml.resl which links worksheets/sheet(i).xml file
    ## order they appear here gives order of worksheets in xlsx file
    sheets$typ <- basename(sheets$Type)
    sheets$target <- stringi::stri_join(xmlDir, xl_path, sheets$Target)
    sheets$id <- as.numeric(sheets$sheetId)

    if (is.null(sheets$state)) sheets$state <- "visible"
    is_visible <- sheets$state %in% c("", "true", "visible")

    ## add worksheets to wb
    # TODO only loop over import_sheets
    for (i in seq_len(nrow(sheets))) {
      if (sheets$typ[i] == "chartsheet") {
        wb$add_chartsheet(sheet = sheets$name[i], visible = is_visible[i])
      } else if (sheets$typ[i] == "worksheet") {
        content_type <- read_xml(ContentTypesXML)
        override <- xml_attr(content_type, "Types", "Override")
        overrideAttr <- as.data.frame(do.call("rbind", override), stringsAsFactors = FALSE)
        xmls <- basename(unlist(overrideAttr$PartName))
        drawings <- grep("drawing", xmls, value = TRUE)
        wb$add_worksheet(sheets$name[i], visible = is_visible[i], has_drawing = !is.na(drawings[i]))
      }
    }

    ## replace sheetId
    for (i in seq_len(nSheets)) {
      wb$workbook$sheets[[i]] <- gsub(
        sprintf(' sheetId="%s"', i),
        sprintf(' sheetId="%s"', sheets$sheetId[i]),
        wb$workbook$sheets[[i]]
      )
    }

    ## additional workbook attributes
    revisionPtr <- xml_node(workbook_xml, "workbook", "xr:revisionPtr")
    if (!data_only && length(revisionPtr)) {
      wb$workbook$revisionPtr <- revisionPtr
    }

    # no clue what calcPr does. If a calcChain is available, this prevents
    # formulas from getting reevaluated unless they are visited manually.
    calcPr <- xml_node(workbook_xml, "workbook", "calcPr")
    if (!data_only && calc_chain && length(calcPr)) {
      # we override the default unless explicitly requested
      if (!(getOption("openxlsx2.disableFullCalcOnLoad", default = FALSE))) {
        calcPr <- xml_attr_mod(calcPr, c(fullCalcOnLoad = "1"))
      }

      wb$workbook$calcPr <- calcPr
    }

    extLst <- xml_node(workbook_xml, "workbook", "extLst")
    if (!data_only && length(extLst)) {

      # remove GoogleSheets "metadata" binary blob
      ext <- xml_node(extLst, "extLst", "ext")
      sel <- which(grepl("GoogleSheets", ext))

      if (length(sel))
        extLst <- xml_rm_child(extLst, "ext", which = sel)

      if (length(xml_node_name(extLst, "extLst")))
        wb$workbook$extLst <- extLst
    }

    workbookPr <- xml_node(workbook_xml, "workbook", "workbookPr")
    if (length(workbookPr)) { # needed for date1904 detection
      wb$workbook$workbookPr <- workbookPr
    }

    workbookProtection <- xml_node(workbook_xml, "workbook", "workbookProtection")
    if (!data_only && length(workbookProtection)) {
      wb$workbook$workbookProtection <- workbookProtection
    }

    customWorkbookViews <- xml_node(workbook_xml, "workbook", "customWorkbookViews")
    if (!data_only && length(customWorkbookViews)) {
      wb$workbook$customWorkbookViews <- customWorkbookViews
    }

    smartTagPr <- xml_node(workbook_xml, "workbook", "smartTagPr")
    if (!data_only && length(smartTagPr)) {
      wb$workbook$smartTagPr <- smartTagPr
    }

    smartTagTypes <- xml_node(workbook_xml, "workbook", "smartTagTypes")
    if (!data_only && length(smartTagTypes)) {
      wb$workbook$smartTagTypes <- smartTagTypes
    }

    externalReferences <- xml_node(workbook_xml, "workbook", "externalReferences")
    if (!data_only && length(externalReferences)) {
      wb$workbook$externalReferences <- externalReferences
    }

    fileRecoveryPr <- xml_node(workbook_xml, "workbook", "fileRecoveryPr")
    if (!data_only && length(fileRecoveryPr)) {
      wb$workbook$fileRecoveryPr <- fileRecoveryPr
    }

    fileSharing <- xml_node(workbook_xml, "workbook", "fileSharing")
    if (!data_only && length(fileSharing)) {
      wb$workbook$fileSharing <- fileSharing
    }

    functionGroups <- xml_node(workbook_xml, "workbook", "functionGroups")
    if (!data_only && length(functionGroups)) {
      wb$workbook$functionGroups <- functionGroups
    }

    oleSize <- xml_node(workbook_xml, "workbook", "oleSize")
    if (!data_only && length(oleSize)) {
      wb$workbook$oleSize <- oleSize
    }

    webPublishing <- xml_node(workbook_xml, "workbook", "webPublishing")
    if (!data_only && length(webPublishing)) {
      wb$workbook$webPublishing <- webPublishing
    }

    webPublishObjects <- xml_node(workbook_xml, "workbook", "webPublishObjects")
    if (!data_only && length(webPublishObjects)) {
      wb$workbook$webPublishObjects <- webPublishObjects
    }

    ## defined Names
    wb$workbook$definedNames <-  xml_node(workbook_xml, "workbook", "definedNames", "definedName")


    ## xti
    if (length(workbookBIN))
      wb$workbook$xti <-  xml_node(workbook_xml, "workbook", "xtis", "xti")

  }

  if (!data_only && calc_chain && length(calcChainXML)) {
    wb$calcChain <- read_xml(calcChainXML, pointer = FALSE)
    wb$append(
      "Content_Types",
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
    )

    ## workbook rels
    wb$append(
      "workbook.xml.rels",
      sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>',
              length(wb$workbook.xml.rels) + 1L
      )
    )
  }

  ## xl\sharedStrings
  if (length(sharedStringsXML)) {

    sst <- read_xml(sharedStringsXML, escapes = TRUE)
    sst_attr <- xml_attr(sst, "sst")
    uniqueCount <- as.character(sst_attr[[1]]["uniqueCount"])
    vals <- xml_node(sst, "sst", "si")
    text <- xml_si_to_txt(sst)

    attr(vals, "uniqueCount") <- uniqueCount
    attr(vals, "text") <- text
    wb$sharedStrings <- vals
  }


  ## xl\metadata
  if (!data_only && length(metadataXML)) {
    wb$append(
      "Content_Types",
      '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>'
    )
    metadata <- read_xml(metadataXML, pointer = FALSE)
    wb$metadata <- metadata
  }

  ## xl\pivotTables & xl\pivotCache
  if (!data_only && length(pivotTableXML)) {

    # pivotTable cacheId links to workbook.xml which links to workbook.xml.rels via rId
    # we don't modify the cacheId, only the rId
    nPivotTables <- length(pivotDefXML)
    rIds <- 20000L + seq_len(nPivotTables)

    wb$pivotDefinitionsRels <- character(nPivotTables)

    pivot_content_type <- NULL

    if (length(pivotTableXML)) {
      wb$pivotTables <- unapply(pivotTableXML, read_xml, pointer = FALSE)
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotTables/pivotTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>', seq_along(pivotTableXML))
      )
    }

    if (length(pivotTableRelsXML)) {
      wb$pivotTables.xml.rels <- unapply(pivotTableRelsXML, read_xml, pointer = FALSE)
    }

    if (length(pivotDefXML)) {
      wb$pivotDefinitions <- unapply(pivotDefXML, read_xml, pointer = FALSE)
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotCache/pivotCacheDefinition%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>', seq_along(pivotDefXML))
      )
    }

    if (length(pivotDefRelsXML)) {
      wb$pivotDefinitionsRels <- unapply(pivotDefRelsXML, read_xml, pointer = FALSE)
    }

    if (length(pivotCacheRecords)) {
      wb$pivotRecords <- unapply(pivotCacheRecords, read_xml, pointer = FALSE)
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotCache/pivotCacheRecords%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>', seq_along(pivotCacheRecords))
      )
    }

    ## update content_types
    wb$append("Content_Types", pivot_content_type)

    ## workbook rels
    wb$workbook.xml.rels <- c(
      wb$workbook.xml.rels,
      sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition%s.xml"/>', rIds, seq_along(pivotDefXML))
    )

    caches <- xml_node(workbook_xml, "workbook", "pivotCaches", "pivotCache")

    for (i in seq_along(caches)) {
      caches[i] <- gsub('"rId[0-9]+"', sprintf('"rId%s"', rIds[i]), caches[i])
    }

    wb$workbook$pivotCaches <- paste0("<pivotCaches>", paste(caches, collapse = ""), "</pivotCaches>")
  }

  ## xl\vbaProject
  if (!data_only && length(vbaProject)) {
    wb$vbaProject <- vbaProject
    wb$Content_Types[grepl('<Override PartName="/xl/workbook.xml" ', wb$Content_Types)] <- '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>'
    wb$append("Content_Types", '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>')
  }


  ## xl\styles
  if (length(stylesXML)) {
    # system(sprintf("cat %s", stylesXML))
    # system(sprintf("cp %s /tmp/styles.xml", stylesXML))
    styles_xml <- read_xml(stylesXML, pointer = FALSE)
    wb$styles_mgr$styles <- import_styles(styles_xml)
    wb$styles_mgr$initialize(wb)
  }

  ## xl\media
  if (!data_only && length(media)) {
    mediaNames <- regmatches(media, regexpr("image[0-9]+\\.[a-z]+$", media))
    fileTypes <- unique(gsub("image[0-9]+\\.", "", mediaNames))

    contentNodes <- sprintf('<Default Extension="%s" ContentType="image/%s"/>', fileTypes, fileTypes)
    contentNodes[fileTypes == "emf"] <- '<Default Extension="emf" ContentType="image/x-emf"/>'

    wb$Content_Types <- c(contentNodes, wb$Content_Types)
    names(media) <- mediaNames
    wb$media <- media
  }

  ## xl\chart
  if (!data_only && (length(chartsXML)) || length(chartExsXML)) {

    # Not every chart has chart, color, style and rel. We read the file names
    # into charts first and replace the file name with the content in a second
    # run.

    # There are some newer charts (presumably all x14), that are written as
    # chartEX and ofc they are counted starting at 1. So the total number of
    # charts is chartsXML + chartsExXML.

    chart_num <- length(chartsXML) + length(chartExsXML)
    empty_chr <- vector("character", chart_num)
    charts <- data.frame(
      chart   = empty_chr,
      chartEx = empty_chr,
      colors  = empty_chr,
      style   = empty_chr,
      rels    = empty_chr,
      relsEx  = empty_chr,
      stringsAsFactors = FALSE
    )

    chartsXML_id        <- filename_id(chartsXML)
    chartExsXML_id      <- filename_id(chartExsXML)
    chartsXML_colors_id <- filename_id(chartsXML_colors)
    chartsXML_styles_id <- filename_id(chartsXML_styles)
    chartsRels_id       <- filename_id(chartsRels)
    chartExsRels_id     <- filename_id(chartExsRels)

    charts$chart[chartsXML_id]         <- names(chartsXML_id)
    charts$chartEx[chartExsXML_id]     <- names(chartExsXML_id)
    charts$colors[chartsXML_colors_id] <- names(chartsXML_colors_id)
    charts$style[chartsXML_styles_id]  <- names(chartsXML_styles_id)
    charts$rels[chartsRels_id]         <- names(chartsRels_id)
    charts$relsEx[chartExsRels_id]     <- names(chartExsRels_id)

    crt_ch <- charts$chart != ""
    crt_ex <- charts$chartEx != ""
    crt_rl <- charts$rels != ""
    crt_re <- charts$relsEx != ""
    crt_co <- charts$colors != ""
    crt_st <- charts$style != ""

    charts$chart[crt_ch]   <- read_xml_files(charts$chart[crt_ch])
    charts$chartEx[crt_ex] <- read_xml_files(charts$chartEx[crt_ex])
    charts$colors[crt_co]  <- read_xml_files(charts$colors[crt_co])
    charts$style[crt_st]   <- read_xml_files(charts$style[crt_st])
    charts$rels[crt_rl]    <- read_xml_files(charts$rels[crt_rl])
    charts$relsEx[crt_re]  <- read_xml_files(charts$relsEx[crt_re])

    wb$charts <- charts

  }

  ## xl\theme
  if (!data_only && length(themeXML)) {
    wb$theme <- read_xml(themeXML, pointer = FALSE)
  }

  ## externalLinks
  if (!data_only && length(extLinksXML)) {
    wb$externalLinks <- lapply(extLinksXML, read_xml, pointer = FALSE)

    wb$append(
      "Content_Types",
      sprintf('<Override PartName="/xl/externalLinks/externalLink%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>', seq_along(extLinksXML))
    )

    ext_ref <- rbindlist(xml_attr(wb$workbook$externalReferences, "externalReferences", "externalReference"))

    for (i in seq_along(extLinksXML)) {
      wb$workbook.xml.rels <- c(
        wb$workbook.xml.rels,
        sprintf(
          '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink%s.xml"/>',
          ext_ref[i, 1],
          i
        )
      )
    }
  }

  ## externalLinksRels
  if (!data_only && length(extLinksRelsXML)) {
    wb$externalLinksRels <- lapply(extLinksRelsXML, read_xml, pointer = FALSE)
  }

  ## featurePropertyBag
  if (!data_only && length(featureProperty)) {
    wb$append(
      "Content_Types",
      '<Override PartName="/xl/featurePropertyBag/featurePropertyBag.xml" ContentType="application/vnd.ms-excel.featurepropertybag+xml"/>'
    )
    wb$featurePropertyBag <- read_xml(featureProperty, pointer = FALSE)
  }


  ##* ----------------------------------------------------------------------------------------------*##
  ### BEGIN READING IN WORKSHEET DATA
  ##* ----------------------------------------------------------------------------------------------*##

  ## xl\worksheets
  file_names <- basename(sheets$Target)

  # nSheets contains all sheets. worksheets and chartsheets. For this loop we
  # only need worksheets. We can not loop over import_sheets, because some
  # might be chart sheets. If a certain sheet is requested, we have to respect
  # this and select only this sheet.

  import_sheets <- seq_len(nrow(sheets))
  if (!missing(sheet)) {
    if (is.null(sheet)) {
      import_sheets <- NULL
    } else {
      import_sheets <- wb_validate_sheet(wb, sheet)
    }

    sheet <- import_sheets
    if (!is.null(sheet) && is.na(sheet)) {
      stop("No such sheet in the workbook. Workbook contains:\n",
           paste(names(wb$get_sheet_names(escape = TRUE)), collapse = "\n"))
    }
  }

  for (i in import_sheets) {

    if (grepl(".bin$", sheets$target[i])) {
      xml_tmp <- gsub(".bin$", ".xml$", sheets$target[i])

      # message(sheets$target[i])
      worksheet_bin(sheets$target[i], wb$is_chartsheet[i], xml_tmp, debug)
      # system(sprintf("cat %s", xml_tmp))
      # system(sprintf("cp %s /tmp/ws.xml", xml_tmp))
      sheets$target[i] <- xml_tmp
    }

    if (sheets$typ[i] == "chartsheet") {
      if (data_only) stop("Requested sheet is a chartsheet. No data to return")
      chartsheet_xml <- read_xml(sheets$target[i])
      wb$worksheets[[i]]$sheetPr <- xml_node(chartsheet_xml, "chartsheet", "sheetPr")
      wb$worksheets[[i]]$sheetViews    <- xml_node(chartsheet_xml, "chartsheet", "sheetViews")
      wb$worksheets[[i]]$sheetProtection    <- xml_node(chartsheet_xml, "chartsheet", "sheetProtection")
      wb$worksheets[[i]]$customSheetViews    <- xml_node(chartsheet_xml, "chartsheet", "customSheetViews")
      wb$worksheets[[i]]$pageMargins <- xml_node(chartsheet_xml, "chartsheet", "pageMargins")
      wb$worksheets[[i]]$pageSetup <- xml_node(chartsheet_xml, "chartsheet", "pageSetup")
      wb$worksheets[[i]]$headerFooter <- xml_node(chartsheet_xml, "chartsheet", "headerFooter")
      wb$worksheets[[i]]$drawing <- xml_node(chartsheet_xml, "chartsheet", "drawing")
      wb$worksheets[[i]]$drawingHF <- xml_node(chartsheet_xml, "chartsheet", "drawingHF")
      wb$worksheets[[i]]$picture <- xml_node(chartsheet_xml, "chartsheet", "picture")
      wb$worksheets[[i]]$webPublishItems <- xml_node(chartsheet_xml, "chartsheet", "webPublishItems")
    } else {
      worksheet_xml <- read_xml(sheets$target[i])
      if (!data_only) {
        wb$worksheets[[i]]$autoFilter <- xml_node(worksheet_xml, "worksheet", "autoFilter")
        wb$worksheets[[i]]$cellWatches <- xml_node(worksheet_xml, "worksheet", "cellWatches")
        wb$worksheets[[i]]$colBreaks <- xml_node(worksheet_xml, "worksheet", "colBreaks", "brk")
        # wb$worksheets[[i]]$cols <- xml_node(worksheet_xml, "worksheet", "cols")
        # wb$worksheets[[i]]$conditionalFormatting <- xml_node(worksheet_xml, "worksheet", "conditionalFormatting")
        wb$worksheets[[i]]$controls <- xml_node(worksheet_xml, "worksheet", "controls")
        wb$worksheets[[i]]$customProperties <- xml_node(worksheet_xml, "worksheet", "customProperties")
        wb$worksheets[[i]]$customSheetViews <- xml_node(worksheet_xml, "worksheet", "customSheetViews")
        wb$worksheets[[i]]$dataConsolidate <- xml_node(worksheet_xml, "worksheet", "dataConsolidate")
        # wb$worksheets[[i]]$dataValidations <- xml_node(worksheet_xml, "worksheet", "dataValidations")
        # wb$worksheets[[i]]$dimension <- xml_node(worksheet_xml, "worksheet", "dimension")
        # has <drawing> a child <legacyDrawing> ?
        wb$worksheets[[i]]$drawing <- xml_node(worksheet_xml, "worksheet", "drawing")
        wb$worksheets[[i]]$drawingHF <- xml_node(worksheet_xml, "worksheet", "drawingHF")
        wb$worksheets[[i]]$legacyDrawing <- xml_node(worksheet_xml, "worksheet", "legacyDrawing")
        wb$worksheets[[i]]$legacyDrawingHF <- xml_node(worksheet_xml, "worksheet", "legacyDrawingHF")
        # wb$worksheets[[i]]$extLst <- xml_node(worksheet_xml, "worksheet", "extLst")
        wb$worksheets[[i]]$headerFooter <- xml_node(worksheet_xml, "worksheet", "headerFooter")
        # wb$worksheets[[i]]$hyperlinks <- xml_node(worksheet_xml, "worksheet", "hyperlinks")
        wb$worksheets[[i]]$ignoredErrors <- xml_node(worksheet_xml, "worksheet", "ignoredErrors")
        # wb$worksheets[[i]]$mergeCells <- xml_node(worksheet_xml, "worksheet", "mergeCells")
        wb$worksheets[[i]]$oleObjects <- xml_node(worksheet_xml, "worksheet", "oleObjects")
        wb$worksheets[[i]]$pageMargins <- xml_node(worksheet_xml, "worksheet", "pageMargins")
        wb$worksheets[[i]]$pageSetup <- xml_node(worksheet_xml, "worksheet", "pageSetup")
        wb$worksheets[[i]]$phoneticPr <- xml_node(worksheet_xml, "worksheet", "phoneticPr")
        wb$worksheets[[i]]$picture <- xml_node(worksheet_xml, "worksheet", "picture")
        wb$worksheets[[i]]$printOptions <- xml_node(worksheet_xml, "worksheet", "printOptions")
        wb$worksheets[[i]]$protectedRanges <- xml_node(worksheet_xml, "worksheet", "protectedRanges")
        wb$worksheets[[i]]$rowBreaks <- xml_node(worksheet_xml, "worksheet", "rowBreaks", "brk")
        wb$worksheets[[i]]$scenarios <- xml_node(worksheet_xml, "worksheet", "scenarios")
        wb$worksheets[[i]]$sheetCalcPr <- xml_node(worksheet_xml, "worksheet", "sheetCalcPr")
        # wb$worksheets[[i]]$sheetData <- xml_node(worksheet_xml, "worksheet", "sheetData")
        # wb$worksheets[[i]]$sheetFormatPr <- xml_node(worksheet_xml, "worksheet", "sheetFormatPr")
        wb$worksheets[[i]]$sheetPr <- xml_node(worksheet_xml, "worksheet", "sheetPr")
        wb$worksheets[[i]]$sheetProtection <- xml_node(worksheet_xml, "worksheet", "sheetProtection")
        # wb$worksheets[[i]]$sheetViews <- xml_node(worksheet_xml, "worksheet", "sheetViews")
        wb$worksheets[[i]]$smartTags <- xml_node(worksheet_xml, "worksheet", "smartTags")
        wb$worksheets[[i]]$sortState <- xml_node(worksheet_xml, "worksheet", "sortState")
        # wb$worksheets[[i]]$tableParts <- xml_node(worksheet_xml, "worksheet", "tableParts")
        wb$worksheets[[i]]$webPublishItems <- xml_node(worksheet_xml, "worksheet", "webPublishItems")

        # check if file loaded was written with incorrect baseColWidth by openxlsx2 <= 1.13
        sheetFormatPr <- xml_node(worksheet_xml, "worksheet", "sheetFormatPr")
        if (length(sheetFormatPr) && sheetFormatPr == "<sheetFormatPr baseColWidth=\"8.43\" defaultRowHeight=\"16\" x14ac:dyDescent=\"0.2\"/>")
          sheetFormatPr <- "<sheetFormatPr baseColWidth=\"8\" defaultRowHeight=\"16\" x14ac:dyDescent=\"0.2\"/>"
        wb$worksheets[[i]]$sheetFormatPr <- sheetFormatPr

        # extract freezePane from sheetViews. This intends to match our freeze
        # pane approach. Though I do not really like it. This blindly shovels
        # everything into sheetViews and freezePane.
        sheetViews    <- xml_node(worksheet_xml, "worksheet", "sheetViews")

        if (length(sheetViews)) {

          # extract only in single sheetView case
          if (length(xml_node(sheetViews, "sheetViews", "sheetView")) == 1) {
            xml_nams <- xml_node_name(sheetViews, "sheetViews", "sheetView")

            freezePane <- paste0(
              vapply(
                unique(xml_nams),
                function(x) {
                  xml <- xml_node(sheetViews, "sheetViews", "sheetView", x)
                  paste(xml, collapse = "")
                },
                NA_character_
              ),
              collapse = ""
            )

            for (xml_nam in xml_nams) {
              sheetViews <- xml_rm_child(
                sheetViews,
                xml_child = xml_nam,
                level = "sheetView"
              )
            }

            wb$worksheets[[i]]$freezePane <- freezePane
          }

        }
        wb$worksheets[[i]]$sheetViews <- sheetViews

        wb$worksheets[[i]]$dataValidations <- xml_node(worksheet_xml, "worksheet", "dataValidations", "dataValidation")
        wb$worksheets[[i]]$extLst <- xml_node(worksheet_xml, "worksheet", "extLst", "ext")
        wb$worksheets[[i]]$tableParts <- xml_node(worksheet_xml, "worksheet", "tableParts", "tablePart")
        wb$worksheets[[i]]$hyperlinks <- xml_node(worksheet_xml, "worksheet", "hyperlinks", "hyperlink")

        # need to expand the names. multiple conditions can be combined in one conditionalFormatting
        cfs <- xml_node(worksheet_xml, "worksheet", "conditionalFormatting")
        if (length(cfs)) {
          nms <- unlist(xml_attr(cfs, "conditionalFormatting"))
          cf <- lapply(cfs, function(x) xml_node(x, "conditionalFormatting", "cfRule"))
          names(cf) <- nms
          conditionalFormatting <- un_list(cf)
          wb$worksheets[[i]]$conditionalFormatting <- conditionalFormatting
        }

      } ## end !data_only

      wb$worksheets[[i]]$dimension  <- xml_node(worksheet_xml, "worksheet", "dimension")
      wb$worksheets[[i]]$cols_attr  <- xml_node(worksheet_xml, "worksheet", "cols", "col")
      wb$worksheets[[i]]$mergeCells <- xml_node(worksheet_xml, "worksheet", "mergeCells", "mergeCell")

      # load the data. This function reads sheet_data and returns cc and row_attr
      loadvals(wb$worksheets[[i]]$sheet_data, worksheet_xml)
    }
  }

  ## Fix headers/footers
  # TODO think about improving headerFooter
  for (i in seq_len(nSheets)) {
    if (sheets$typ[i] == "worksheet") {
      if (length(wb$worksheets[[i]]$headerFooter)) {
        # get attributes
        hf_attrs <- rbindlist(xml_attr(wb$worksheets[[i]]$headerFooter, "headerFooter"))

        if (!is.null(hf_attrs$alignWithMargins))
          wb$worksheets[[i]]$align_with_margins <- as.logical(as.integer(hf_attrs$alignWithMargins))

        if (!is.null(hf_attrs$scaleWithDoc))
          wb$worksheets[[i]]$scale_with_doc     <- as.logical(as.integer(hf_attrs$scaleWithDoc))

        wb$worksheets[[i]]$headerFooter       <- getHeaderFooterNode(wb$worksheets[[i]]$headerFooter)
      }
    }
  }


  ##* ----------------------------------------------------------------------------------------------*##
  ### READING IN WORKSHEET DATA COMPLETE
  ##* ----------------------------------------------------------------------------------------------*##


  ## Next sheetRels to see which drawings_rels belongs to which sheet
  if (!data_only && length(sheetRelsXML)) {

    ## sheetrId is order sheet appears in xlsx file
    ## create a 1-1 vector of rels to worksheet
    ## haveRels is boolean vector where i-the element is TRUE/FALSE if sheet has a rels sheet

    if (length(chartSheetsXML) == 0) {
      allRels <- file.path(dirname(sheetRelsXML[1]), paste0(file_names, ".rels"))
      haveRels <- allRels %in% sheetRelsXML
    } else {
      haveRels <- rep(FALSE, length(wb$worksheets))
      allRels <- rep("", length(wb$worksheets))

      for (i in seq_len(nSheets)) {
        if (sheets$typ[i] == "chartsheet") {
          rels_file <- file.path(xmlDir, "xl", "chartsheets", "_rels", paste0(file_names[i], ".rels"))
        } else {
          rels_file <- file.path(xmlDir, "xl", "worksheets", "_rels", paste0(file_names[i], ".rels"))
        }
        if (file.exists(rels_file)) {
          allRels[i] <- rels_file
          haveRels[i] <- TRUE
        }
      }
    }

    ## sheet.xml have been reordered to be in the order of sheetrId
    ## not every sheet has a worksheet rels

    xml <- lapply(seq_along(allRels), function(i) {
      if (haveRels[i]) {
        xml <- xml_node(allRels[[i]], "Relationships", "Relationship")
        if (length(xml) == 0) return(character())

        xml_relship <- rbindlist(xml_attr(xml, "Relationship"))
        # print(xml_relship)
        if (any(basename(xml_relship$Type) %in% c("comments", "table"))) { #  %% length(tablesBIN)
          sel <- basename(xml_relship$Type) %in% c("comments", "table")
          # message("table")
          # print(gsub(".bin", ".xml", xml_relship$Target[sel]))
          # message("---")
          xml_relship$Target[sel] <- gsub(".bin", ".xml", xml_relship$Target[sel])
        }

        if (is.null(xml_relship$TargetMode)) xml_relship$TargetMode <- ""

        # we do not ship this binary blob, therefore spreadsheet software may
        # stumble over this non existent reference. In the future we might want
        # to check if the references are valid pre file saving.
        sel_row <- !grepl("printerSettings|binaryIndex", basename2(xml_relship$Target))
        sel_col <- c("Id", "Type", "Target", "TargetMode")
        # return as xml
        xml <- df_to_xml("Relationship", xml_relship[sel_row, sel_col])
      } else {
        xml <- character()
      }
      xml
    })

    wb$worksheets_rels <- xml

    for (ws in seq_along(wb$worksheets)) {

      # This relships tracks the file numbering. drawing1.xml or comments2.xml
      # hyperlinks are not in files, but xml links in the references
      wb_rels <- rbindlist(xml_attr(wb$worksheets_rels[[ws]], "Relationship"))
      cmmts <- integer()
      drwns <- integer()
      hyper <- integer()
      pvtbl <- integer()
      slcrs <- integer()
      table <- integer()
      trcmt <- integer()
      tmlne <- integer()
      vmldr <- integer()

      if (ncol(wb_rels)) {
        # since target can be any hyperlink, we have to expect various things here like uint64
        wb_rels$tid <- suppressWarnings(as.integer(gsub("\\D+", "", basename2(wb_rels$Target))))
        wb_rels$typ <- basename(wb_rels$Type)

        # for hyperlinks, we take the relationship id
        if (length(wb_rels$typ == "hyperlink")) {
          wb_rels$tid[wb_rels$typ == "hyperlink"] <- as.integer(
            gsub("\\D+", "", basename2(wb_rels$Id[wb_rels$typ == "hyperlink"]))
          )
        }

        cmmts <- wb_rels$tid[wb_rels$typ == "comments"]
        drwns <- wb_rels$tid[wb_rels$typ == "drawing"]
        hyper <- wb_rels$tid[wb_rels$typ == "hyperlink"]
        pvtbl <- wb_rels$tid[wb_rels$typ == "pivotTable"]
        slcrs <- wb_rels$tid[wb_rels$typ == "slicer"]
        table <- wb_rels$tid[wb_rels$typ == "table"]
        trcmt <- wb_rels$tid[wb_rels$typ == "threadedComment"]
        tmlne <- wb_rels$tid[wb_rels$typ == "timeline"]
        vmldr <- wb_rels$tid[wb_rels$typ == "vmlDrawing"]
      }

      # currently we use only a selected set of these
      # as of 0.5.9000: we use comments, drawing, and vmlDrawing
      wb$worksheets[[ws]]$relships <- list(
        comments         = cmmts,
        drawing          = drwns,
        hyperlink        = hyper,
        pivotTable       = pvtbl,
        slicer           = slcrs,
        table            = table,
        threadedComment  = trcmt,
        timeline         = tmlne,
        vmlDrawing       = vmldr
      )
    }


    ## Slicers -------------------------------------------------------------------------------------
    if (length(slicerXML)) {

      # maybe these need to be sorted?
      # slicerXML <- slicerXML[order(nchar(slicerXML), slicerXML)] ???

      wb$slicers <- vapply(slicerXML, read_xml, pointer = FALSE,
                           FUN.VALUE = NA_character_, USE.NAMES = FALSE)

      ## worksheet_rels Id for slicer will be rId0
      for (i in seq_along(wb$slicers)) {

        # this will add slicers to Content_Types. Ergo if worksheets with
        # slicers are removed, the slicer needs to remain in the worksheet
        wb$append(
          "Content_Types",
          sprintf('<Override PartName="/xl/slicers/slicer%s.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>', i)
        )
      }
    }

    ## ---- slicerCaches
    if (length(slicerCachesXML)) {
      wb$slicerCaches <- vapply(slicerCachesXML, read_xml, pointer = FALSE,
                                FUN.VALUE = NA_character_, USE.NAMES = FALSE)

      for (i in seq_along(wb$slicerCaches)) {
        wb$append("Content_Types", sprintf('<Override PartName="/xl/slicerCaches/slicerCache%s.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>', i))
        wb$append("workbook.xml.rels", sprintf('<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache%s.xml"/>', 1E5 + i, i))
      }

      # get extLst object. select the slicerCaches and replace it
      ext_nams <- xml_node_name(wb$workbook$extLst, "extLst", "ext")
      is_slicer <- which(ext_nams %in% c("x14:slicerCaches", "x15:slicerCaches"))
      ext <- xml_node(wb$workbook$extLst, "extLst", "ext")
      ext[is_slicer] <- genSlicerCachesExtLst(1E5 + seq_along(slicerCachesXML))
      wb$workbook$extLst <- xml_node_create("extLst", xml_children = ext)
    }

    ## Timeline -------------------------------------------------------------------------------------
    if (length(timelineXML)) {

      # maybe these need to be sorted?
      # timelineXML <- timelineXML[order(nchar(timelineXML), timelineXML)] ???

      wb$timelines <- vapply(timelineXML, read_xml, pointer = FALSE,
                             FUN.VALUE = NA_character_, USE.NAMES = FALSE)

      ## worksheet_rels Id for timeline will be rId0
      for (i in seq_along(wb$timelines)) {

        # this will add timelines to Content_Types. Ergo if worksheets with
        # timelines are removed, the timeline needs to remain in the worksheet
        wb$append(
          "Content_Types",
          sprintf('<Override PartName="/xl/timelines/timeline%s.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>', i)
        )
      }
    }

    ## ---- timelineCaches
    if (length(timelineCachesXML)) {
      wb$timelineCaches <- vapply(timelineCachesXML, read_xml, pointer = FALSE,
                                  FUN.VALUE = NA_character_, USE.NAMES = FALSE)

      for (i in seq_along(wb$timelineCaches)) {
        wb$append("Content_Types", sprintf('<Override PartName="/xl/timelineCaches/timelineCache%s.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>', i))
        wb$append("workbook.xml.rels", sprintf('<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache" Target="timelineCaches/timelineCache%s.xml"/>', 2E5 + i, i))
      }

      # get extLst object. select the timelineCaches and replace it
      ext_nams <- xml_node_name(wb$workbook$extLst, "extLst", "ext")
      is_timeline <- which(ext_nams == "x15:timelineCacheRefs")
      ext <- xml_node(wb$workbook$extLst, "extLst", "ext")
      ext[is_timeline] <- genTimelineCachesExtLst(2E5 + seq_along(timelineCachesXML))
      wb$workbook$extLst <- xml_node_create("extLst", xml_children = ext)
    }

    ## Tables --------------------------------------------------------------------------------------
    if (length(tablesXML)) {
      tables <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=table)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      tableSheets <- unapply(seq_along(sheets$`r:id`), function(i) rep(i, length(tables[[i]])))

      ## sort the tables into the order they appear in the xml and tables variables
      names(tablesXML) <- basename(tablesXML)
      tablesXML <- tablesXML[sprintf("table%s.xml", unlist(tables))]
      ## tables are now in correct order so we can read them in as they are

      tables_xml <- vapply(tablesXML, FUN = read_xml, pointer = FALSE, FUN.VALUE = NA_character_)
      tabs <- rbindlist(xml_attr(tables_xml, "table"))

      wb$tables <- data.frame(
        tab_name = tabs[["displayName"]],
        tab_sheet = tableSheets,
        tab_ref = tabs[["ref"]],
        tab_xml = as.character(tables_xml),
        tab_act = 1,
        stringsAsFactors = FALSE
      )

      wb$append("Content_Types", sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', nrow(wb$tables)))

      ## every worksheet containing a table, has a table part. this references
      ## the display name, so that we know which tables are still around.
      for (i in seq_along(tableSheets)) {
        table_sheet_i <- tableSheets[i]
        attr(wb$worksheets[[table_sheet_i]]$tableParts, "tableName") <- c(attr(wb$worksheets[[table_sheet_i]]$tableParts, "tableName"), tabs[["displayName"]][i])
      }

    } ## if (length(tablesXML))


    ## Drawings ------------------------------------------------------------------------------------
    if (length(drawingsXML)) {

      drw_len <- max(as.integer(gsub("\\D+", "", basename(drawingsXML))))

      wb$drawings      <- rep(list(""), drw_len) # vector("list", drw_len)
      wb$drawings_rels <- rep(list(""), drw_len) # vector("list", drw_len)


      for (drw in drawingsXML) {

        drw_file <- as.integer(gsub("\\D+", "", basename(drw)))

        wb$drawings[drw_file] <- read_xml(drw, pointer = FALSE)
      }

      for (drw_rel in drawingRelsXML) {

        drw_file <- as.integer(gsub("\\D+", "", basename(drw_rel)))

        wb$drawings_rels[[drw_file]] <- xml_node(drw_rel, "Relationships", "Relationship")
      }

    }


    ## VML Drawings --------------------------------------------------------------------------------
    if (length(vmlDrawingXML)) {
      wb$append("Content_Types", '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>')

      vml_len <- max(as.integer(gsub("\\D+", "", basename(vmlDrawingXML))))

      wb$vml      <- rep(list(""), vml_len) # vector("list", vml_len)
      wb$vml_rels <- rep(list(""), vml_len) # vector("list", vml_len)


      for (vml in vmlDrawingXML) {

        vml_file <- as.integer(gsub("\\D+", "", basename(vml)))

        # fix broken xml in vml buttons
        vml <- stringi::stri_read_lines(vml, encoding = "UTF-8")
        vml <- paste(vml, sep = "", collapse = "")
        vml <- gsub("<br>(?!</br>)", "<br/>", vml, perl = TRUE)
        wb$vml[vml_file] <- read_xml(vml, pointer = FALSE)
      }

      for (vml_rel in vmlDrawingRelsXML) {

        vml_file <- as.integer(gsub("\\D+", "", basename(vml_rel)))

        wb$vml_rels[[vml_file]] <- xml_node(vml_rel, "Relationships", "Relationship")
      }

    }

    # remove drawings from Content_Types. These drawings are the old imported drawings.
    # we will add drawings only when writing and will use the sheet to create them.
    wb$Content_Types <- wb$Content_Types[!grepl("drawings/drawing", wb$Content_Types)]

    ## vmlDrawing and comments
    if (length(commentsXML)) {

      wb$comments <- vector("list", length(commentsXML))

      for (comment_xml in seq_along(commentsXML)) {
        # read xml and split into authors and comments
        txt <- read_xml(commentsXML[comment_xml])
        authors <- xml_value(txt, "comments", "authors", "author")
        comments <- xml_node(txt, "comments", "commentList", "comment")


      # create valid rich text strings in comments table
      if (length(commentsBIN) && any(sel <- grepl("<FONT_\\d+/>", comments))) {

        SST      <- c(comments)

        matches <- stringi::stri_extract_all_regex(SST[sel], "<FONT_\\d+/>")
        matches <- unique(unlist(matches))

        values  <- as.integer(gsub("\\D+", "", matches))

        xmls    <- stringi::stri_replace_all_fixed(
          wb$styles_mgr$styles$fonts[values + 1],
          c("<name", "font>"),
          c("<rFont", "rPr>"),
          vectorize_all = FALSE
        )

        sst <- stringi::stri_replace_all_fixed(
          str           = SST[sel],
          pattern       = matches,
          replacement   = xmls,
          vectorize_all = FALSE
        )

        SST[sel] <- sst

        comments <- SST
      }

        comments_attr <- rbindlist(xml_attr(comments, "comment"))

        refs <- comments_attr$ref
        authorsInds <- as.integer(comments_attr$authorId) + 1
        authors <- authors[authorsInds]

        text <- xml_node(comments, "comment", "text")

        comments <- lapply(comments, function(x) {
          text <- xml_node(x, "comment", "text")
          if (all(xml_node_name(x, "comment", "text") == "t")) {
            list(
              style = FALSE,
              comments = xml_node(text, "text", "t")
            )
          } else {
            list(
              style = xml_node(text, "text", "r", "rPr"),
              comments = xml_node(text, "text", "r", "t")
            )
          }
        })

        wb$comments[[comment_xml]] <- lapply(seq_along(comments), function(j) {
          list(
            #"refId" = com_rId[j],
            "ref" = refs[j],
            "author" = authors[j],
            "comment" = comments[[j]]$comments,
            "style" = comments[[j]]$style
          )
        })
      }
    }

    ## Threaded comments
    if (any(length(threadCommentsXML) > 0)) {

      if (any(lengths(threadCommentsXML))) {
        wb$threadComments <- lapply(threadCommentsXML, function(x) xml_node(x, "ThreadedComments", "threadedComment"))
      }

      wb$append(
        "Content_Types",
        sprintf('<Override PartName="/xl/threadedComments/%s" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>',
                vapply(threadCommentsXML, basename, NA_character_))
      )
    }

    ## Persons (needed for Threaded Comment)
    if (length(personXML) > 0) {
      wb$persons <- read_xml(personXML, pointer = FALSE)

      wb$append(
        "Content_Types",
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>'
      )
      wb$append(
        "workbook.xml.rels",
        '<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>')
    }

    ## Embedded docx
    if (length(embeddings)) {

      # get the embedded files extensions
      files <- unique(gsub(".+\\.(\\w+)$", "\\1", embeddings))

      # get the required ContentTypes
      content_type <- read_xml(ContentTypesXML)
      extensions <- rbindlist(xml_attr(content_type, "Types", "Default"))
      extensions <- extensions[extensions$Extension %in% files, ]

      # append the content types
      default <- sprintf('<Default Extension="%s" ContentType="%s"/>',
                         extensions$Extension, extensions$ContentType)
      wb$append("Content_Types", default)

      wb$embeddings <- embeddings
    }

    ## xl\activeX
    if (length(activeX)) {

      wb$activeX <- activeX
      ax_sel <- file_ext2(activeX) == "xml"
      ax_fls <- basename2(activeX[ax_sel])

      wb$append("Content_Types", '<Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/>')
      wb$append("Content_Types", sprintf('<Override PartName="/xl/activeX/%s" ContentType="application/vnd.ms-office.activeX+xml"/>', ax_fls))
    }

  } else {
    # If workbook contains no sheetRels, create empty workbook.xml.rels.
    # Otherwise spreadsheet software will stumble over missing rels to drwaing.
    wb$worksheets_rels <- lapply(seq_along(wb$sheet_names), FUN = function(x) character())
  } ## end of worksheetRels

  ## queryTables
  if (!data_only && length(queryTablesXML) > 0) {
    ids <- as.numeric(regmatches(queryTablesXML, regexpr("[0-9]+(?=\\.xml)", queryTablesXML, perl = TRUE)))
    wb$queryTables <- unapply(queryTablesXML[order(ids)], read_xml, pointer = FALSE)
    wb$append(
      "Content_Types",
      sprintf('<Override PartName="/xl/queryTables/queryTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>', seq_along(queryTablesXML))
    )
  }


  ## connections
  if (!data_only && length(connectionsXML) > 0) {
    wb$connections <- read_xml(connectionsXML, pointer = FALSE)
    wb$append("workbook.xml.rels", '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>')
    wb$append("Content_Types", '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>')
  }

  # In files with data querry connections we have a folder called "customXml" at the top file level
  # We probably will never modify this folder. Therefore we copy it when saving and apped entries to
  # Content_Types and workbook.xml.rels. The actual rId does not really seem to matter.
  if (!data_only && length(customXmlDir)) {

    wb$customXml <- customXmlDir

    for (cstxml in seq_along(grep_xml("/customXml/itemProps"))) {
      wb$append("Content_Types",
                sprintf('<Override PartName="/customXml/itemProps%s.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>',
                        cstxml)
      )
    }

    for (cstitm in seq_along(grep_xml("customXml/item[0-9]+.xml"))) {

      # TODO provide a function that creates a wb_rels data frame
      wb_rels <- rbindlist(xml_attr(wb$workbook.xml.rels, "Relationship"))
      wb_rels$typ <- basename(wb_rels$Type)
      wb_rels$id  <- as.numeric(gsub("\\D", "", wb_rels$Id))
      next_rid <- max(wb_rels$id) + 1

      wb$append("workbook.xml.rels",
                sprintf(
                  '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item%s.xml"/>',
                  next_rid,
                  cstitm
                )
      )
    }
  }

  if (!data_only && length(ctrlPropsXML)) {
    wb$ctrlProps <- read_xml_files(ctrlPropsXML)
    for (ctrlProp in seq_along(ctrlPropsXML)) {
      wb$append("Content_Types",
                sprintf(
                  '<Override PartName="/xl/ctrlProps/ctrlProp%s.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>',
                  ctrlProp
                )
      )
    }
  }


  ## table rels
  if (!data_only && length(tableRelsXML)) {

    ## table_i_might have tableRels_i but I am re-ordering the tables to be in order of worksheets
    ## I make every table have a table_rels so i need to fill in the gaps if any table_rels are missing

    tmp <- paste0(basename(tablesXML), ".rels")
    hasRels <- tmp %in% basename(tableRelsXML)

    ## order tableRelsXML
    tableRelsXML <- tableRelsXML[match(tmp[hasRels], basename(tableRelsXML))]

    ##
    wb$tables.xml.rels <- character(length = length(tablesXML))

    ## which sheet does it belong to
    xml <- sapply(tableRelsXML, read_xml, pointer = FALSE)

    wb$tables.xml.rels[hasRels] <- xml
  } else if (length(tablesXML) > 0) {
    wb$tables.xml.rels <- rep("", length(tablesXML))
  }

  # xlsb files have references as index position stored. we have replaced the
  # index position with "openxlsx2xlsb_" + indexpos. We have exported the xti
  # references as custom xml node in the workbook. Now we have to create the
  # correct sheet references and replace our replacement with it.
  if (!data_only && length(workbookBIN)) {

    # we need to update the order of customSheetView children. Incorrect orders
    # causes spreadsheet software to be unable to load and recover the file.
    for (sheet in seq_along(wb$worksheets)) {
      if (length(wb$worksheets[[sheet]]$customSheetViews) == 0) next

      cvs <- xml_node(wb$worksheets[[sheet]]$customSheetViews, "customSheetViews", "customSheetView")

       # chart sheets have a reduced custom view
      exp_attr <- c(
        "guid", "scale", "colorId", "showPageBreaks", "showFormulas",
        "showGridLines", "showRowCol", "outlineSymbols", "zeroValues",
        "fitToPage", "printArea", "filter", "showAutoFilter", "hiddenRows",
        "hiddenColumns", "state", "filterUnique", "view", "showRuler",
        "topLeftCell", "zoomToFit"
      )
      exp_nams <- c(
        "pane", "selection", "rowBreaks", "colBreaks", "pageMargins",
        "printOptions", "pageSetup", "headerFooter", "autoFilter", "extLst"
      )
      cv <- read_xml2df(read_xml(cvs), "customSheetView", vec_attrs = exp_attr, vec_chlds = exp_nams)

      # headerFooter cause issues. they are (a) not added to the correct node
      # and (b) brick the entire XML structure
      cv$headerFooter <- ""

      cvs <- write_df2xml(cv[c(exp_attr, exp_nams)], "customSheetView", vec_attrs = exp_attr, vec_chlds = exp_nams)

      wb$worksheets[[sheet]]$customSheetViews <- xml_node_create("customSheetViews", xml_children = cvs)
    }

    if (length(wb$workbook$xti)) {
      # create data frame containing sheet names for Xti entries
      xti <- rbindlist(xml_attr(wb$workbook$xti, "xti"))
      xti$name_id <- paste0("openxlsx2xlsb_", sprintf("%012d", as.integer(rownames(xti)) - 1L))
      xti$firstSheet <- as.integer(xti$firstSheet) + 1L
      xti$lastSheet <- as.integer(xti$lastSheet) + 1L

      sheets <- wb$get_sheet_names(escape = TRUE)

      xti$sheets <- "" # was NA_character_ but (in missing cases: all is <NA>)
      # all id == 0 are local references, otherwise external references
      # external references are written as "[0]sheetname!A1". Require
      # handling of externalReferences.bin
      if (debug) print(xti)

      sel <- !grepl("^rId", xti$type) & xti$firstSheet >= 0

      if (any(sel)) {

          for (i in seq_len(nrow(xti[sel, ]))) {
            want <- xti$firstSheet[sel][i]

            # want can be zero
            if (want %in% seq_along(sheets)) {

              sheetName <- sheets[[want]]
              if (xti$firstSheet[sel][i] < xti$lastSheet[sel][i]) {
                want <- xti$lastSheet[sel][i]
                sheetName <- paste0(sheetName, ":", sheets[[want]])
              }

              # should be a single reference now
              val <- sheetName

              if (grepl("[^A-Za-z0-9]", val))
                val <- shQuote(val, type = "sh")

              if (length(val)) xti$sheets[sel][i] <- val
            }
          }
      }

      for (i in seq_along(wb$tables$tab_xml)) {
        wb$tables$tab_xml[i] <-
          stringi::stri_replace_all_fixed(
            wb$tables$tab_xml[i],
            xti$name_id,
            xti$sheets,
            vectorize_all = FALSE
          )
      }

      for (i in seq_along(wb$worksheets)) {
        if (!wb$is_chartsheet[[i]])
          wb$worksheets[[i]]$extLst <-
            stringi::stri_replace_all_fixed(
              wb$worksheets[[i]]$extLst,
              xti$name_id,
              xti$sheets,
              vectorize_all = FALSE
            )
      }

      ### for external references we need to get the required sheet names first
      # For now this is all a little guess work

      # This will return all sheets of the external reference.
      extSheets <- lapply(wb$externalLinks, function(x) {
        sheetNames <- xml_node(x, "externalLink", "externalBook", "sheetNames")
        # assuming that all external links have some kind of vals
        rbindlist(xml_attr(sheetNames, "sheetNames", "sheetName"))[["val"]]
      })

      # TODO: How are references to the same external link,
      # but different sheet handled?
      sel <- grepl("^rId", xti$type) & xti$firstSheet >= 0

      if (any(sel)) {

        xti$ext_id <- NA_integer_
        xti$ext_id[sel] <- as.integer(as.factor(as.integer(gsub("\\D+", "", xti$type[sel]))))

        # loop over it and create external link
        for (i in seq_len(nrow(xti[sel, ]))) {
          want <- xti$firstSheet[sel][i]
          ref  <- xti$ext_id[sel][i]

          # want can be zero
          if (ref %in% seq_along(extSheets)) {

            sheetName <- extSheets[[ref]][[want]]
            if (xti$firstSheet[sel][i] < xti$lastSheet[sel][i]) {
              want <- xti$lastSheet[sel][i]
              sheetName <- paste0(sheetName, ":", extSheets[[ref]][[want]])
            }
            # should be a single reference now

            val <- sprintf("[%s]%s", xti$ext_id[sel][i], sheetName)

            # non ascii or whitespace
            if (grepl("[^A-Za-z0-9]", val))
              val <- shQuote(val, type = "sh")

            if (length(val)) xti$sheets[sel][i] <- val
          }
        }
      }

      if (debug)
        print(xti)

      if (nrow(xti)) {

        if (length(wb$workbook$definedNames)) {

          wb$workbook$definedNames <-
            stringi::stri_replace_all_fixed(
              wb$workbook$definedNames,
              xti$name_id,
              xti$sheets,
              vectorize_all = FALSE
            )

          # replace named region in formulas
          nri         <- wb$get_named_regions()
          nri$name_id <- paste0("openxlsx2defnam_", sprintf("%012d", as.integer(nri$id)))

          if (debug)
            print(nri)

          for (j in seq_along(wb$worksheets)) {
            if (any(sel <- grepl(paste0(nri$name_id, collapse = "|"), wb$worksheets[[j]]$sheet_data$cc$f))) {
              wb$worksheets[[j]]$sheet_data$cc$f[sel] <-
                stringi::stri_replace_all_fixed(
                  wb$worksheets[[j]]$sheet_data$cc$f[sel],
                  nri$name_id,
                  nri$name,
                  vectorize_all = FALSE
                )
            }


            if (!wb$is_chartsheet[[i]])
              wb$worksheets[[j]]$dataValidations <-
                stringi::stri_replace_all_fixed(
                  wb$worksheets[[j]]$dataValidations,
                  nri$name_id,
                  nri$name,
                  vectorize_all = FALSE
                )
          }

        }

        if (length(wb$tables)) {
          # replace named region in formulas
          tri         <- wb$get_tables(sheet = NULL)
          tri$id      <- as.integer(rbindlist(xml_attr(wb$tables$tab_xml, "table"))$id) # - 1L
          tri$name_id <- paste0("openxlsx2tab_", sprintf("%012d", tri$id))
          tri$vars    <- lapply(wb$tables$tab_xml, function(x) rbindlist(xml_attr(x, "table", "tableColumns", "tableColumn"))$name)

          tri <- tri[order(tri$id), ]

          if (debug)
            print(tri)

          for (j in seq_along(wb$worksheets)) {
            if (any(sel <- grepl(paste0(tri$name_id, collapse = "|"), wb$worksheets[[j]]$sheet_data$cc$f))) {

              for (i in seq_len(nrow(tri))) {

                sel <- grepl(paste0(tri$name_id, collapse = "|"), wb$worksheets[[j]]$sheet_data$cc$f)

                from_xlsb <- c(tri$name_id[i], paste0("[openxlsx2col_", tri$id[i], "_", seq_along(unlist(tri$vars[i])) - 1L, "]"))
                to_xlsx   <- c(tri$tab_name[i], paste0("[", unlist(tri$vars[i]), "]"))

                # always on all?
                wb$tables$tab_xml <-
                  stringi::stri_replace_all_fixed(
                    wb$tables$tab_xml,
                    from_xlsb,
                    to_xlsx,
                    vectorize_all = FALSE
                  )

                wb$worksheets[[j]]$sheet_data$cc$f[sel] <-
                  stringi::stri_replace_all_fixed(
                    wb$worksheets[[j]]$sheet_data$cc$f[sel],
                    from_xlsb,
                    to_xlsx,
                    vectorize_all = FALSE
                  )

                wb$workbook$definedNames <-
                  stringi::stri_replace_all_fixed(
                    wb$workbook$definedNames,
                    from_xlsb,
                    to_xlsx,
                    vectorize_all = FALSE
                  )
              }
            }
          }

        }

        # this might be terribly slow!
        for (j in seq_along(wb$worksheets)) {
          if (any(sel <- wb$worksheets[[j]]$sheet_data$cc$f != "")) {
            wb$worksheets[[j]]$sheet_data$cc$f[sel] <-
              stringi::stri_replace_all_fixed(
                wb$worksheets[[j]]$sheet_data$cc$f[sel],
                xti$name_id,
                xti$sheets,
                vectorize_all = FALSE
              )
          }
        }
      }
    }

    # remove defined names without value. these are not valid and are probably
    # remnants of xti
    if (length(dfn_nms <- wb$workbook$definedNames)) {
      wb$workbook$definedNames <- dfn_nms[xml_value(dfn_nms, "definedName") != ""]
    }

    # create valid rich text strings in shared strings table
    if (any(sel <- grepl("<FONT_\\d+/>", wb$sharedStrings))) {

      attr_sst <- attributes(wb$sharedStrings)
      SST      <- c(wb$sharedStrings)

      matches <- stringi::stri_extract_all_regex(SST[sel], "<FONT_\\d+/>")
      matches <- unique(unlist(matches))

      values  <- as.integer(gsub("\\D+", "", matches))

      xmls    <- stringi::stri_replace_all_fixed(
        wb$styles_mgr$styles$fonts[values + 1],
        c("<name", "font>"),
        c("<rFont", "rPr>"),
        vectorize_all = FALSE
      )

      sst <- stringi::stri_replace_all_fixed(
        str           = SST[sel],
        pattern       = matches,
        replacement   = xmls,
        vectorize_all = FALSE
      )

      SST[sel] <- sst
      attributes(SST) <- attr_sst

      wb$sharedStrings <- SST
    }

  }

  ## richData ------------------------------------------------------------------------------------
  # This is new in openxlsx2 1.6 and probably not yet entire correct
  if (!data_only && (length(richValueRel) || length(rdrichvalue) || length(rdrichvaluestr) || length(rdRichValueTypes))) {

    rd <- data.frame(
      richValueRel     = "",
      richValueRelrels = "",
      rdrichvalue      = "",
      rdrichvaluestr   = "",
      rdRichValueTypes = "",
      stringsAsFactors = FALSE
    )

    if (length(richValueRel)) {
      wb$append(
        "Content_Types",
        '<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>'
      )
      wb$append(
        "workbook.xml.rels",
        '<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>'
      )
      rd$richValueRel <- read_xml(richValueRel, pointer = FALSE)
    }

    if (length(richValueRelrels)) {
      rd$richValueRelrels <- read_xml(richValueRelrels, pointer = FALSE)
    }

    if (length(rdrichvalue)) {
      wb$append(
        "Content_Types",
        '<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>'
      )
      wb$append(
        "workbook.xml.rels",
        '<Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>'
      )
      rd$rdrichvalue <- read_xml(rdrichvalue, pointer = FALSE)
    }

    if (length(rdrichvaluestr)) {
      wb$append(
        "Content_Types",
        '<Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>'
      )
      wb$append(
        "workbook.xml.rels",
        '<Relationship Id="rId7" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>'
      )
      rd$rdrichvaluestr <- read_xml(rdrichvaluestr, pointer = FALSE)
    }

    if (length(rdRichValueTypes)) {
      wb$append(
        "Content_Types",
        '<Override PartName="/xl/richData/rdRichValueTypes.xml" ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>'
      )
      wb$append(
        "workbook.xml.rels",
        '<Relationship Id="rId8" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes" Target="richData/rdRichValueTypes.xml"/>'
      )
      rd$rdRichValueTypes <- read_xml(rdRichValueTypes, pointer = FALSE)
    }

    wb$richData <- rd
  }


  # final cleanup
  if (length(workbookBIN)) {
    wb$workbook$xti <- NULL
  }


  return(wb)
}
