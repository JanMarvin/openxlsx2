#' @name wb_load
#' @title Load an existing .xlsx file
#' @param file A path to an existing .xlsx or .xlsm file
#' @param xlsxFile alias for file
#' @param sheet optional sheet parameter. if this is applied, only the selected
#' sheet will be loaded.
#' @param data_only mode to import if only a data frame should be returned. This
#' strips the wbWorkbook to a bare minimum.
#' @param calc_chain optionally you can keep the calculation chain intact. This
#' is used by spreadsheet software to identify the order in which formulas are
#' evaluated. Removing the calculation chain is considered harmless. The calc
#' chain will be created upon the next time the worksheet is loaded in
#' spreadsheet software. Keeping it, might only speed loading time in said
#' software.
#' @description  wb_load returns a workbook object conserving styles and
#' formatting of the original .xlsx file.
#' @details A warning is displayed if an xml namespace for main is found in the
#' xlsx file. Certain xlsx files created by third-party applications contain a
#' namespace (usually `x`). This namespace is not required for the file to work
#' in spreadsheet software and is not expected by `openxlsx2`. Therefore it is
#' removed when the file is loaded into a workbook. Removal is generally
#' expected to be safe, but the feature is still experimental.
#' @return Workbook object.
#' @export
#' @seealso [wb_remove_worksheet()]
#' @examples
#' ## load existing workbook from package folder
#' wb <- wb_load(file = system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2"))
#' wb$get_sheet_names() # list worksheets
#' wb ## view object
#' ## Add a worksheet
#' wb$add_worksheet("A new worksheet")
wb_load <- function(
    file,
    xlsxFile = NULL,
    sheet,
    data_only = FALSE,
    calc_chain = FALSE
) {

  file <- xlsxFile %||% file
  file <- getFile(file)

  if (!file.exists(file)) {
    stop("File does not exist.")
  }

  ## create temp dir
  xmlDir <- temp_dir("_openxlsx_wb_load")

  # do not unlink after loading
  # on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)

  ## Unzip files to temp directory
  xmlFiles <- unzip(file, exdir = xmlDir)
  # we need to read the files in human order: 1, 2, 10 and not 1, 10, 2.
  ordr <- stringi::stri_order(xmlFiles, opts_collator = stringi::stri_opts_collator(numeric = TRUE))
  xmlFiles <- xmlFiles[ordr]

  wb <- wb_workbook()

  grep_xml <- function(pattern, perl = TRUE, value = TRUE, ...) {
    # targets xmlFiles; has presents
    grep(pattern, xmlFiles, perl = perl, value = value, ...)
  }

  ## Not used
  # .relsXML          <- grep_xml("_rels/.rels$")
  ContentTypesXML   <- grep_xml("\\[Content_Types\\].xml$")
  appXML            <- grep_xml("app.xml$")
  coreXML           <- grep_xml("core.xml$")
  customXML         <- grep_xml("custom.xml$")

  customXmlDir      <- grep_xml("customXml/")

  workbookXML       <- grep_xml("workbook.xml$")
  workbookXMLRels   <- grep_xml("workbook.xml.rels")

  drawingsXML       <- grep_xml("drawings/drawing[0-9]+.xml$")
  worksheetsXML     <- grep_xml("/worksheets/sheet[0-9]+")

  stylesXML         <- grep_xml("styles.xml$")
  sharedStringsXML  <- grep_xml("sharedStrings.xml$")
  metadataXML       <- grep_xml("metadata.xml$")
  themeXML          <- grep_xml("theme[0-9]+.xml$")
  drawingRelsXML    <- grep_xml("drawing[0-9]+.xml.rels$")
  sheetRelsXML      <- grep_xml("sheet[0-9]+.xml.rels$")
  media             <- grep_xml("image[0-9]+.[a-z]+$")
  vmlDrawingXML     <- grep_xml("drawings/vmlDrawing[0-9]+\\.vml$")
  vmlDrawingRelsXML <- grep_xml("vmlDrawing[0-9]+.vml.rels$")
  calcChainXML      <- grep_xml("xl/calcChain.xml")
  embeddings        <- grep_xml("xl/embeddings")

  # comments
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
  chartSheetsXML    <- grep_xml("xl/chartsheets/sheet[0-9]+\\.xml")

  # tables
  tablesXML         <- grep_xml("tables/table[0-9]+.xml$")
  tableRelsXML      <- grep_xml("table[0-9]+.xml.rels$")
  queryTablesXML    <- grep_xml("queryTable[0-9]+.xml$")

  # connections
  connectionsXML    <- grep_xml("connections.xml$")
  extLinksXML       <- grep_xml("externalLink[0-9]+.xml$")
  extLinksRelsXML   <- grep_xml("externalLink[0-9]+.xml.rels$")

  # form control
  ctrlPropsXML      <- grep_xml("ctrlProps/ctrlProp[0-9]+.xml")

  # pivot tables
  pivotTableXML     <- grep_xml("pivotTable[0-9]+.xml$")
  pivotTableRelsXML <- grep_xml("pivotTable[0-9]+.xml.rels$")
  pivotDefXML       <- grep_xml("pivotCacheDefinition[0-9]+.xml$")
  pivotDefRelsXML   <- grep_xml("pivotCacheDefinition[0-9]+.xml.rels$")
  pivotCacheRecords <- grep_xml("pivotCacheRecords[0-9]+.xml$")

  ## slicers
  slicerXML         <- grep_xml("slicer[0-9]+.xml$")
  slicerCachesXML   <- grep_xml("slicerCache[0-9]+.xml$")

  ## VBA Macro
  vbaProject        <- grep_xml("vbaProject\\.bin$")

  ## remove all EXCEPT media and charts
  on.exit(
    unlink(
      # TODO: this removes all files, the folders remain. grep instead grep_xml?
      grep_xml("media|vmlDrawing|customXml|comment|embeddings|vbaProject|person", ignore.case = TRUE, invert = TRUE),
      recursive = TRUE, force = TRUE
    ),
    add = TRUE
  )

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
  chartSheetRIds <- NULL
  if (!data_only && length(chartSheetsXML)) {
    workbookRelsXML <- grep("chartsheets/sheet", workbookRelsXML, fixed = TRUE, value = TRUE)

    chartSheetRIds <- unlist(getId(workbookRelsXML))
    chartsheet_rId_mapping <- unlist(regmatches(workbookRelsXML, gregexpr("sheet[0-9]+\\.xml", workbookRelsXML, perl = TRUE, ignore.case = TRUE)))

    sheetNo <- as.integer(regmatches(chartSheetsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml)", chartSheetsXML, perl = TRUE)))
    chartSheetsXML <- chartSheetsXML[order(sheetNo)]

    chartSheetsRelsXML <- grep_xml("xl/chartsheets/_rels")
    sheetNo2 <- as.integer(regmatches(chartSheetsRelsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml\\.rels)", chartSheetsRelsXML, perl = TRUE)))
    chartSheetsRelsXML <- chartSheetsRelsXML[order(sheetNo2)]

    chartSheetsRelsDir <- dirname(chartSheetsRelsXML[1])
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

    ## Some veryHidden sheets do not have a sheet content and their rId is empty.
    ## Such sheets need to be filtered out because otherwise their sheet names
    ## occur in the list of all sheet names, leading to a wrong association
    ## of sheet names with sheet indeces.
    sheets <- sheets[sheets$`r:id` != "", ]

    # if wb_relsxml is not available, the workbook has no relationships, not
    # sure if this is possible
    if (length(workbookXMLRels)) {
      wb_rels_xml <- rbindlist(
        xml_attr(workbookXMLRels, "Relationships", "Relationship")
      )
    }

    sheets <- merge(
      sheets, wb_rels_xml,
      by.x = "r:id", by.y = "Id",
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
    sheets$target <- stri_join(xmlDir, xl_path, sheets$Target)
    sheets$id <- as.numeric(sheets$sheetId)

    if (is.null(sheets$state)) sheets$state <- "visible"
    is_visible <- sheets$state %in% c("", "true", "visible")

    ## add worksheets to wb
    # TODO only loop over import_sheets
    for (i in seq_len(nrow(sheets))) {
      if (sheets$typ[i] == "chartsheet") {
        txt <- read_xml(sheets$target[i], pointer = FALSE)
        wb$add_chartsheet(sheet = sheets$name[i], visible = is_visible[i])
      } else if (sheets$typ[i] == "worksheet") {
        content_type <- read_xml(ContentTypesXML)
        override <- xml_attr(content_type, "Types", "Override")
        overrideAttr <- as.data.frame(do.call("rbind", override))
        xmls <- basename(unlist(overrideAttr$PartName))
        drawings <- grep("drawing", xmls, value = TRUE)
        wb$add_worksheet(sheets$name[i], visible = is_visible[i], hasDrawing = !is.na(drawings[i]))
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
      wb$workbook$extLst <- extLst
    }

    workbookPr <- xml_node(workbook_xml, "workbook", "workbookPr")
    if (!data_only && length(workbookPr)) {
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


  ## xl\sharedStrings
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
      relsEx  = empty_chr
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
    import_sheets <- wb_validate_sheet(wb, sheet)
    sheet <- import_sheets
    if (is.na(sheet)) {
      stop("No such sheet in the workbook. Workbook contains:\n",
           paste(names(wb$get_sheet_names()), collapse = "\n"))
    }
  }

  for (i in import_sheets) {
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

        wb$worksheets[[i]]$sheetFormatPr <- xml_node(worksheet_xml, "worksheet", "sheetFormatPr")

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

        wb$worksheets[[i]]$cols_attr     <- xml_node(worksheet_xml, "worksheet", "cols", "col")

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

      wb$worksheets[[i]]$dimension <- xml_node(worksheet_xml, "worksheet", "dimension")
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
        wb$worksheets[[i]]$headerFooter <- getHeaderFooterNode(wb$worksheets[[i]]$headerFooter)
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
          ind <- which(chartSheetRIds == sheets$`r:id`[i])
          rels_file <- file.path(chartSheetsRelsDir, paste0(chartsheet_rId_mapping[ind], ".rels"))
        } else {
          ind <- sheets$`r:id`[i]
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
        # xml_relship$Target[basename(xml_relship$Type) == "drawing"] <- sprintf("../drawings/drawing%s.xml", i)
        # xml_relship$Target[basename(xml_relship$Type) == "vmlDrawing"] <- sprintf("../drawings/vmlDrawing%s.vml", i)

        if (is.null(xml_relship$TargetMode)) xml_relship$TargetMode <- ""

        # we do not ship this binary blob, therefore spreadsheet software may
        # stumble over this non existent reference. In the future we might want
        # to check if the references are valid pre file saving.
        sel_row <- !grepl("printerSettings", basename(xml_relship$Target))
        sel_col <- c("Id", "Type", "Target", "TargetMode")
        # return as xml
        xml <- df_to_xml("Relationship", xml_relship[sel_row, sel_col])
      } else {
        xml <- character()
      }
      return(xml)
    })

    wb$worksheets_rels <- xml

    xml <- lapply(seq_along(allRels), function(i) {
      if (haveRels[i]) {
        xml <- xml_node(allRels[[i]], "Relationships", "Relationship")
      } else {
        xml <- character()
      }
      return(xml)
    })

    for (ws in seq_along(wb$worksheets)) {

      # This relships tracks the file numbering. drawing1.xml or comments2.xml
      # hyperlinks are not in files, but xml links in the references
      wb_rels <- rbindlist(xml_attr(wb$worksheets_rels[[ws]], "Relationship"))
      cmmts <- integer()
      drwns <- integer()
      pvtbl <- integer()
      slcrs <- integer()
      table <- integer()
      trcmt <- integer()
      vmldr <- integer()

      if (ncol(wb_rels)) {
        # since target can be any hyperlink, we have to expect various things here like uint64
        wb_rels$tid <- suppressWarnings(as.integer(gsub("\\D+", "", basename(wb_rels$Target))))
        wb_rels$typ <- basename(wb_rels$Type)

        cmmts <- wb_rels$tid[wb_rels$typ == "comments"]
        drwns <- wb_rels$tid[wb_rels$typ == "drawing"]
        pvtbl <- wb_rels$tid[wb_rels$typ == "pivotTable"]
        slcrs <- wb_rels$tid[wb_rels$typ == "slicer"]
        table <- wb_rels$tid[wb_rels$typ == "table"]
        trcmt <- wb_rels$tid[wb_rels$typ == "threadedComment"]
        vmldr <- wb_rels$tid[wb_rels$typ == "vmlDrawing"]
      }

      # currently we use only a selected set of these
      # as of 0.5.9000: we use comments, drawing, and vmlDrawing
      wb$worksheets[[ws]]$relships <- list(
        comments         = cmmts,
        drawing          = drwns,
        pivotTable       = pvtbl,
        slicer           = slcrs,
        table            = table,
        threadedComment  = trcmt,
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
      is_slicer <- which(ext_nams == "x14:slicerCaches")
      ext <- xml_node(wb$workbook$extLst, "extLst", "ext")
      ext[is_slicer] <- genSlicerCachesExtLst(1E5 + seq_along(slicerCachesXML))
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

      wb$append("Content_Types", sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', nrow(wb$tables)))

      # # TODO When does this happen?
      # if (length(tabs["displayName"]) != length(tablesXML)) {
      #   tabs[["displayName"]] <- paste0("Table", seq_along(tablesXML))
      # }

      wb$tables <- data.frame(
        tab_name = tabs[["displayName"]],
        tab_sheet = tableSheets,
        tab_ref = tabs[["ref"]],
        tab_xml = as.character(tables_xml),
        tab_act = 1,
        stringsAsFactors = FALSE
      )

      # ## relabel ids
      # for (i in seq_len(nrow(wb$tables))) {
      #   wb$tables$tab_xml[i] <- xml_attr_mod(wb$tables$tab_xml[i], xml_attributes = c(id = as.character(i + 2)))
      # }

      ## every worksheet containing a table, has a table part. this references
      ## the display name, so that we know which tables are still around.
      for (i in seq_along(tableSheets)) {
        table_sheet_i <- tableSheets[i]
        attr(wb$worksheets[[table_sheet_i]]$tableParts, "tableName") <- c(attr(wb$worksheets[[table_sheet_i]]$tableParts, "tableName"), tabs[["displayName"]][i])
      }

    } ## if (length(tablesXML))

    ## might we have some external hyperlinks
    # TODO use lengths()
    if (any(vapply(wb$worksheets[sheets$typ == "worksheet"], function(x) length(x$hyperlinks), NA_integer_) > 0)) {

      ## Do we have external hyperlinks
      hlinks <- lapply(xml, function(x) x[grepl("hyperlink", x) & grepl("External", x)])
      # TODO use lengths()
      hlinksInds <- which(lengths(hlinks) > 0)

      ## If it's an external hyperlink it will have a target in the sheet_rels
      if (length(hlinksInds)) {
        for (i in hlinksInds) {
          ids <- apply_reg_match(hlinks[[i]], '(?<=Id=").*?"')
          ids <- gsub('"$', "", ids)

          targets <- apply_reg_match(hlinks[[i]], '(?<=Target=").*?"')
          targets <- gsub('"$', "", targets)

          ids2 <- lapply(wb$worksheets[[i]]$hyperlinks, reg_match, pat = '(?<=r:id=").*?"')
          ids2[lengths(ids2) == 0] <- NA
          ids2 <- gsub('"$', "", unlist(ids2))

          targets <- targets[match(ids2, ids)]
          names(wb$worksheets[[i]]$hyperlinks) <- targets
        }
      }

      # remove unused hyperlink reference from worksheets_rels
      wb$worksheets_rels[[i]] <- relship_no(wb$worksheets_rels[[i]], "hyperlink")
    }


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
        vml <- gsub("<br>", "<br/>", vml)
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

        comments_attr <- rbindlist(xml_attr(comments, "comment"))

        refs <- comments_attr$ref
        authorsInds <- as.integer(comments_attr$authorId) + 1
        authors <- authors[authorsInds]

        text <- xml_node(comments, "comment", "text")

        comments <- lapply(comments, function(x) {
          text <- xml_node(x, "comment", "text")
          list(
            style = xml_node(text, "text", "r", "rPr"),
            comments = xml_node(text, "text", "r", "t")
          )
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
    if (length(threadCommentsXML) > 0) {
      threadCommentsXMLrelationship <- lapply(xml, function(x) grep("threadedComment[0-9]+\\.xml", x, value = TRUE))
      hasThreadComments <- lengths(threadCommentsXMLrelationship) > 0
      if (any(hasThreadComments)) {
        for (i in seq_along(xml)) {
          if (hasThreadComments[i]) {
            target <- apply_reg_match(threadCommentsXMLrelationship[[i]], '(?<=Target=").*?"')
            target <- basename(gsub('"$', "", target))

            wb$threadComments[[i]] <- grep(target, threadCommentsXML, value = TRUE)

          }
        }
      }
      wb$append(
        "Content_Types",
        sprintf('<Override PartName="/xl/threadedComments/%s" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>',
                vapply(threadCommentsXML, basename, NA_character_))
      )
    }

    ## Persons (needed for Threaded Comment)
    if (length(personXML) > 0) {
      wb$persons <- personXML
      wb$append(
        "Content_Types",
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>'
      )
      wb$append(
        "workbook.xml.rels",
        '<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>')
    }

    ## Embedded docx
    if (length(embeddings) > 0) {
      # TODO only valid for docx. need to check xls and doc?
      wb$append("Content_Types", '<Default Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"/>')
      wb$embeddings <- embeddings
    }

  } else {
    # If workbook contains no sheetRels, create empty workbook.xml.rels.
    # Otherwise spreadsheet software will stumble over missing rels to drwaing.
    wb$worksheets_rels <- lapply(seq_along(wb$sheet_names), FUN = function(x) character())
  } ## end of worksheetRels

  ## convert hyperliks to hyperlink objects
  if (!data_only)
    for (i in seq_len(nSheets)) {
      if (!wb$is_chartsheet[i])
        wb$worksheets[[i]]$hyperlinks <- xml_to_hyperlink(wb$worksheets[[i]]$hyperlinks)
    }

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

  return(wb)
}
