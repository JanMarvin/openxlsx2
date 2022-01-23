


#' @name loadWorkbook
#' @title Load an existing .xlsx file
#' @param file A path to an existing .xlsx or .xlsm file
#' @param xlsxFile alias for file
#' @param isUnzipped Set to TRUE if the xlsx file is already unzipped
#' @description  loadWorkbook returns a workbook object conserving styles and
#' formatting of the original .xlsx file.
#' @return Workbook object.
#' @export
#' @seealso [removeWorksheet()]
#' @examples
#' ## load existing workbook from package folder
#' wb <- loadWorkbook(file = system.file("extdata", "loadExample.xlsx", package = "openxlsx2"))
#' names(wb) # list worksheets
#' wb ## view object
#' ## Add a worksheet
#' addWorksheet(wb, "A new worksheet")
#'
#' ## Save workbook
#' \dontrun{
#' saveWorkbook(wb, "loadExample.xlsx", overwrite = TRUE)
#' }
#'
loadWorkbook <- function(file, xlsxFile = NULL, isUnzipped = FALSE) {
  # TODO  default of isUnzipped to tools::file_ext(file) == "zip" ?
  ## If this is a unzipped workbook, skip the temp dir stuff
  if (isUnzipped) {
    xmlDir <- file
    xmlFiles <- list.files(path = xmlDir, full.names = TRUE, recursive = TRUE, all.files = TRUE)
  } else {
    if (!is.null(xlsxFile)) {
      file <- xlsxFile
    }

    file <- getFile(file)

    file <- getFile(file)
    if (!file.exists(file)) {
      stop("File does not exist.")
    }

    ## create temp dir
    xmlDir <- file.path(tempdir(), paste0(tempfile(tmpdir = ""), "_openxlsx_loadWorkbook"))

    ## Unzip files to temp directory
    xmlFiles <- unzip(file, exdir = xmlDir)
  }
  wb <- createWorkbook()


  grep_xml <- function(pattern, perl = TRUE, value = TRUE, ...) {
    # targets xmlFiles; has presents
    grep(pattern, xmlFiles, perl = perl, value = value, ...)
  }

  ## Not used
  # .relsXML          <- grep_xml("_rels/.rels$")
  appXML            <- grep_xml("app.xml$")

  ContentTypesXML   <- grep_xml("\\[Content_Types\\].xml$")

  drawingsXML       <- grep_xml("drawings/drawing[0-9]+.xml$")
  worksheetsXML     <- grep_xml("/worksheets/sheet[0-9]+")

  coreXML           <- grep_xml("core.xml$")
  workbookXML       <- grep_xml("workbook.xml$")
  stylesXML         <- grep_xml("styles.xml$")
  sharedStringsXML  <- grep_xml("sharedStrings.xml$")
  themeXML          <- grep_xml("theme[0-9]+.xml$")
  drawingRelsXML    <- grep_xml("drawing[0-9]+.xml.rels$")
  sheetRelsXML      <- grep_xml("sheet[0-9]+.xml.rels$")
  media             <- grep_xml("image[0-9]+.[a-z]+$")
  vmlDrawingXML     <- grep_xml("drawings/vmlDrawing[0-9]+\\.vml$")
  vmlDrawingRelsXML <- grep_xml("vmlDrawing[0-9]+.vml.rels$")
  calcChainXML      <- grep_xml("xl/calcChain.xml")
  commentsXML       <- grep_xml("xl/comments[0-9]+\\.xml")
  threadCommentsXML <- grep_xml("xl/threadedComments/threadedComment[0-9]+\\.xml")
  personXML         <- grep_xml("xl/persons/person.xml$")
  commentsrelXML    <- grep_xml("xl/worksheets/_rels/sheet[0-9]+\\.xml")
  embeddings        <- grep_xml("xl/embeddings")

  charts            <- grep_xml("xl/charts/.*xml$")
  chartsRels        <- grep_xml("xl/charts/_rels")
  chartSheetsXML    <- grep_xml("xl/chartsheets/sheet[0-9]+\\.xml")

  tablesXML         <- grep_xml("tables/table[0-9]+.xml$")
  tableRelsXML      <- grep_xml("table[0-9]+.xml.rels$")
  queryTablesXML    <- grep_xml("queryTable[0-9]+.xml$")
  connectionsXML    <- grep_xml("connections.xml$")
  extLinksXML       <- grep_xml("externalLink[0-9]+.xml$")
  extLinksRelsXML   <- grep_xml("externalLink[0-9]+.xml.rels$")


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
  if (!isUnzipped) {
    on.exit(
      unlink(
        grep_xml("charts|media|vmlDrawing|comment|embeddings|pivot|slicer|vbaProject|person", ignore.case = TRUE, invert = TRUE),
        recursive = TRUE, force = TRUE
      ),
      add = TRUE
    )
  }

  ## core
  if (length(coreXML) == 1) {
    wb$core <- read_xml(coreXML, pointer = FALSE)
  }

  if (length(appXML)) {
    wb$app <- read_xml(appXML, pointer = FALSE)
  }

  nSheets <- length(worksheetsXML) + length(chartSheetsXML)

  ## get Rid of chartsheets, these do not have a worksheet/sheeti.xml
  worksheet_rId_mapping <- NULL
  workbookRelsXML <- grep_xml("workbook.xml.rels$")
  if (length(workbookRelsXML)) {
    xml <- read_xml(workbookRelsXML)
    workbookRelsXML <- xml_node(xml, "Relationships", "Relationship")
    worksheet_rId_mapping <- grep("worksheets/sheet", workbookRelsXML, fixed = TRUE, value = TRUE)
  }

  ##
  chartSheetRIds <- NULL
  if (length(chartSheetsXML)) {
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
    workbook_xml <- read_xml(workbookXML)

    wb$workbook$fileVersion <- xml_node(workbook_xml, "workbook", "fileVersion")
    wb$workbook$alternateContent <- xml_node(workbook_xml, "workbook", "mc:AlternateContent")
    wb$workbook$bookViews <- xml_node(workbook_xml, "workbook", "bookViews")

    sheets <- xml_attribute(workbook_xml, "workbook", "sheets", "sheet")
    sheets <- rbindlist(sheets)

    ## Some veryHidden sheets do not have a sheet content and their rId is empty.
    ## Such sheets need to be filtered out because otherwise their sheet names
    ## occur in the list of all sheet names, leading to a wrong association
    ## of sheet names with sheet indeces.
    sheets <- sheets[sheets$`r:id` != "",]


    ## sheetId is meaningless
    ## sheet rId links to the workbook.xml.resl which links worksheets/sheet(i).xml file
    ## order they appear here gives order of worksheets in xlsx file

    sheetrId <- sheets$`r:id`
    sheetId <- sheets$sheetId
    sheetNames <- sheets$name

    is_chart_sheet <- sheetrId %in% chartSheetRIds
    if (is.null(sheets$state)) sheets$state <- "visible"
    is_visible <- sheets$state == "visible"

    ## add worksheets to wb
    j <- 1
    for (i in seq_along(sheetrId)) {
      if (is_chart_sheet[i]) {
        count <- 0
        txt <- read_xml(chartSheetsXML[j])

        zoom <- regmatches(txt, regexpr('(?<=zoomScale=")[0-9]+', txt, perl = TRUE))
        if (length(zoom) == 0) {
          zoom <- 100
        }

        tabColour <- xml_node(txt, "worksheet", "sheetPr", "tabColor")
        if (length(tabColour) == 0) {
          tabColour <- NULL
        }

        j <- j + 1L

        wb$addChartSheet(sheetName = sheetNames[i], tabColour = tabColour, zoom = as.numeric(zoom))
      } else {
        content_type <- read_xml(ContentTypesXML)
        override <- xml_attribute(content_type, "Types", "Override")
        overrideAttr <- as.data.frame(do.call("rbind", override))
        xmls <- basename(unlist(overrideAttr$PartName))
        drawings <- grep("drawing", xmls, value = TRUE)
        wb$addWorksheet(sheetNames[i], visible = is_visible[i], hasDrawing = !is.na(drawings[i]))
      }
    }


    ## replace sheetId
    for (i in seq_len(nSheets)) {
      wb$workbook$sheets[[i]] <- gsub(sprintf(' sheetId="%s"', i), sprintf(' sheetId="%s"', sheetId[i]), wb$workbook$sheets[[i]])
    }


    ## additional workbook attributes
    calcPr <- xml_node(workbook_xml, "workbook", "calcPr")
    if (length(calcPr)) {
      wb$workbook$calcPr <- calcPr
    }

    workbookPr <- xml_node(workbook_xml, "workbook", "workbookPr")
    if (length(workbookPr)) {
      wb$workbook$workbookPr <- workbookPr
    }

    workbookProtection <- xml_node(workbook_xml, "workbook", "workbookProtection")
    if (length(workbookProtection)) {
      wb$workbook$workbookProtection <- workbookProtection
    }

    customWorkbookViews <- xml_node(workbook_xml, "workbook", "customWorkbookViews")
    if (length(customWorkbookViews)) {
      wb$workbook$customWorkbookViews <- customWorkbookViews
    }

    smartTagPr <- xml_node(workbook_xml, "workbook", "smartTagPr")
    if (length(smartTagPr)) {
      wb$workbook$smartTagPr <- smartTagPr
    }

    smartTagTypes <- xml_node(workbook_xml, "workbook", "smartTagTypes")
    if (length(smartTagTypes)) {
      wb$workbook$smartTagTypes <- smartTagTypes
    }

    webPublishing <- xml_node(workbook_xml, "workbook", "webPublishing")
    if (length(webPublishing)) {
      wb$workbook$webPublishing <- webPublishing
    }

    externalReferences <- xml_node(workbook_xml, "workbook", "externalReferences")
    if (length(externalReferences)) {
      wb$workbook$externalReferences <- externalReferences
    }

    fileRecoveryPr <- xml_node(workbook_xml, "workbook", "fileRecoveryPr")
    if (length(fileRecoveryPr)) {
      wb$workbook$fileRecoveryPr <- fileRecoveryPr
    }

    fileSharing <- xml_node(workbook_xml, "workbook", "fileSharing")
    if (length(fileSharing)) {
      wb$workbook$fileSharing <- fileSharing
    }

    functionGroups <- xml_node(workbook_xml, "workbook", "functionGroups")
    if (length(functionGroups)) {
      wb$workbook$functionGroups <- functionGroups
    }

    oleSize <- xml_node(workbook_xml, "workbook", "oleSize")
    if (length(oleSize)) {
      wb$workbook$oleSize <- oleSize
    }

    webPublishing <- xml_node(workbook_xml, "workbook", "webPublishing")
    if (length(webPublishing)) {
      wb$workbook$webPublishing <- webPublishing
    }

    webPublishObjects <- xml_node(workbook_xml, "workbook", "webPublishObjects")
    if (length(webPublishObjects)) {
      wb$workbook$webPublishObjects <- webPublishObjects
    }

    webPublishObjects <- xml_node(workbook_xml, "workbook", "webPublishObjects")
    if (length(webPublishObjects)) {
      wb$workbook$webPublishObjects <- webPublishObjects
    }

    ## defined Names
    wb$workbook$definedNames <-  xml_node(workbook_xml, "workbook", "definedNames", "definedName")

  }

  if (length(calcChainXML)) {
    wb$calcChain <- read_xml(calcChainXML, pointer = FALSE)
    wb$Content_Types <- c(
      wb$Content_Types,
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
    )
  }



  ## xl\sharedStrings
  if (length(sharedStringsXML)) {

    sst <- read_xml(sharedStringsXML)
    uniqueCount <- getXMLXPtr1attr_one(sst, "sst", "uniqueCount")
    vals <- xml_node(sst, "sst", "si")
    text <- si_to_txt(sst)

    attr(vals, "uniqueCount") <- uniqueCount
    attr(vals, "text") <- text
    wb$sharedStrings <- vals
  }

  ## xl\pivotTables & xl\pivotCache
  if (length(pivotTableXML)) {

    # pivotTable cacheId links to workbook.xml which links to workbook.xml.rels via rId
    # we don't modify the cacheId, only the rId
    nPivotTables <- length(pivotDefXML)
    rIds <- 20000L + seq_len(nPivotTables)

    ## pivot tables
    pivotTableXML <- pivotTableXML[order(nchar(pivotTableXML), pivotTableXML)]
    pivotTableRelsXML <- pivotTableRelsXML[order(nchar(pivotTableRelsXML), pivotTableRelsXML)]

    ## Cache
    pivotDefXML <- pivotDefXML[order(nchar(pivotDefXML), pivotDefXML)]
    pivotDefRelsXML <- pivotDefRelsXML[order(nchar(pivotDefRelsXML), pivotDefRelsXML)]
    pivotCacheRecords <- pivotCacheRecords[order(nchar(pivotCacheRecords), pivotCacheRecords)]


    wb$pivotDefinitionsRels <- character(nPivotTables)

    pivot_content_type <- NULL

    if (length(pivotTableRelsXML)) {
      wb$pivotTables.xml.rels <- unlist(lapply(pivotTableRelsXML, read_xml, pointer = FALSE))
    }


    # ## Check what caches are used
    cache_keep <- unlist(regmatches(wb$pivotTables.xml.rels, gregexpr("(?<=pivotCache/pivotCacheDefinition)[0-9](?=\\.xml)",
      wb$pivotTables.xml.rels,
      perl = TRUE, ignore.case = TRUE
    )))

    ## pivot cache records
    tmp <- unlist(regmatches(pivotCacheRecords, gregexpr("(?<=pivotCache/pivotCacheRecords)[0-9]+(?=\\.xml)", pivotCacheRecords, perl = TRUE, ignore.case = TRUE)))
    pivotCacheRecords <- pivotCacheRecords[tmp %in% cache_keep]

    ## pivot cache definitions rels
    tmp <- unlist(regmatches(pivotDefRelsXML, gregexpr("(?<=_rels/pivotCacheDefinition)[0-9]+(?=\\.xml)", pivotDefRelsXML, perl = TRUE, ignore.case = TRUE)))
    pivotDefRelsXML <- pivotDefRelsXML[tmp %in% cache_keep]

    ## pivot cache definitions
    tmp <- unlist(regmatches(pivotDefXML, gregexpr("(?<=pivotCache/pivotCacheDefinition)[0-9]+(?=\\.xml)", pivotDefXML, perl = TRUE, ignore.case = TRUE)))
    pivotDefXML <- pivotDefXML[tmp %in% cache_keep]



    if (length(pivotTableXML)) {
      wb$pivotTables[seq_along(pivotTableXML)] <- pivotTableXML
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotTables/pivotTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>', seq_along(pivotTableXML))
      )
    }

    if (length(pivotDefXML)) {
      wb$pivotDefinitions[seq_along(pivotDefXML)] <- pivotDefXML
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotCache/pivotCacheDefinition%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>', seq_along(pivotDefXML))
      )
    }

    if (length(pivotCacheRecords)) {
      wb$pivotRecords[seq_along(pivotCacheRecords)] <- pivotCacheRecords
      pivot_content_type <- c(
        pivot_content_type,
        sprintf('<Override PartName="/xl/pivotCache/pivotCacheRecords%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>', seq_along(pivotCacheRecords))
      )
    }

    if (length(pivotDefRelsXML)) {
      wb$pivotDefinitionsRels[seq_along(pivotDefRelsXML)] <- pivotDefRelsXML
    }




    ## update content_types
    wb$Content_Types <- c(wb$Content_Types, pivot_content_type)


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
  if (length(vbaProject)) {
    wb$vbaProject <- vbaProject
    wb$Content_Types[grepl('<Override PartName="/xl/workbook.xml" ', wb$Content_Types)] <- '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>'
    wb$Content_Types <- c(wb$Content_Types, '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>')
  }


  ## xl\styles
  if (length(stylesXML) > 0) {
    # assign("styleObjects", styleObjects, globalenv())
    wb$styles_xml <- read_xml(stylesXML, pointer = FALSE)

    wb$styles <- import_styles(wb$styles_xml)

  } else {
    wb$styleObjects <- list()
  }

  ## xl\media
  if (length(media)) {
    mediaNames <- regmatches(media, regexpr("image[0-9]+\\.[a-z]+$", media))
    fileTypes <- unique(gsub("image[0-9]+\\.", "", mediaNames))

    contentNodes <- sprintf('<Default Extension="%s" ContentType="image/%s"/>', fileTypes, fileTypes)
    contentNodes[fileTypes == "emf"] <- '<Default Extension="emf" ContentType="image/x-emf"/>'

    wb$Content_Types <- c(contentNodes, wb$Content_Types)
    names(media) <- mediaNames
    wb$media <- media
  }



  ## xl\chart
  if (length(charts)) {
    chartNames <- basename(charts)
    nCharts <- sum(grepl("chart[0-9]+.xml", chartNames))
    nChartStyles <- sum(grepl("style[0-9]+.xml", chartNames))
    nChartCol <- sum(grepl("colors[0-9]+.xml", chartNames))

    if (nCharts > 0) {
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/chart%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', seq_len(nCharts)))
    }

    if (nChartStyles > 0) {
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/style%s.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>', seq_len(nChartStyles)))
    }

    if (nChartCol > 0) {
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/colors%s.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>', seq_len(nChartCol)))
    }

    if (length(chartsRels)) {
      charts <- c(charts, chartsRels)
      chartNames <- c(chartNames, file.path("_rels", basename(chartsRels)))
    }

    names(charts) <- chartNames
    wb$charts <- charts
  }






  ## xl\theme
  if (length(themeXML)) {
    wb$theme <- read_xml(themeXML, pointer = FALSE)
  }


  ## externalLinks
  if (length(extLinksXML)) {
    wb$externalLinks <- lapply(sort(extLinksXML), read_xml, pointer = FALSE)

    wb$Content_Types <- c(
      wb$Content_Types,
      sprintf('<Override PartName="/xl/externalLinks/externalLink%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>', seq_along(extLinksXML))
    )

    wb$workbook.xml.rels <- c(wb$workbook.xml.rels, sprintf(
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>',
      seq_along(extLinksXML)
    ))
  }

  ## externalLinksRels
  if (length(extLinksRelsXML)) {
    wb$externalLinksRels <- lapply(sort(extLinksRelsXML), read_xml, pointer = FALSE)
  }







  ##* ----------------------------------------------------------------------------------------------*##
  ### BEGIN READING IN WORKSHEET DATA
  ##* ----------------------------------------------------------------------------------------------*##

  ## xl\worksheets
  file_names <- regmatches(worksheet_rId_mapping, regexpr("sheet[0-9]+\\.xml", worksheet_rId_mapping, perl = TRUE))
  file_rIds <- unlist(getId(worksheet_rId_mapping))
  file_names <- file_names[match(sheetrId, file_rIds)]

  worksheetsXML <- file.path(dirname(worksheetsXML), file_names)


  # TODO this loop should live in loadworksheets
  for (i in seq_len(nSheets)) {
    worksheet_xml <- read_xml(worksheetsXML[i])

    wb$worksheets[[i]]$autoFilter <- xml_node(worksheet_xml, "worksheet", "autoFilter")
    wb$worksheets[[i]]$cellWatches <- xml_node(worksheet_xml, "worksheet", "cellWatches")
    wb$worksheets[[i]]$colBreaks <- xml_node(worksheet_xml, "worksheet", "colBreaks")
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
    wb$worksheets[[i]]$rowBreaks <- xml_node(worksheet_xml, "worksheet", "rowBreaks")
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



    wb$worksheets[[i]]$dimension <- xml_node(worksheet_xml, "worksheet", "dimension")

    wb$worksheets[[i]]$sheetFormatPr <- xml_node(worksheet_xml, "worksheet", "sheetFormatPr")
    wb$worksheets[[i]]$sheetViews    <- xml_node(worksheet_xml, "worksheet", "sheetViews")
    wb$worksheets[[i]]$cols_attr     <- xml_node(worksheet_xml, "worksheet", "cols", "col")


    # need to expand the names. multiple conditions can be combined in one conditionalFormatting
    cfs <- xml_node(worksheet_xml, "worksheet", "conditionalFormatting")
    if (length(cfs)) {
      nms <- unlist(xml_attribute(cfs, "conditionalFormatting"))
      cf <- lapply(cfs, function(x) xml_node(x, "conditionalFormatting", "cfRule"))
      names(cf) <- nms
      conditionalFormatting <- unlist(cf)
      names(conditionalFormatting) <- unlist(lapply(nms, function(x) rep(x, length(cf[[x]]))))

      wb$worksheets[[i]]$conditionalFormatting <- conditionalFormatting
    }
    wb$worksheets[[i]]$sheetProtection <- xml_node(worksheet_xml, "worksheet", "sheetProtection")

    wb$worksheets[[i]]$dataValidations <- xml_node(worksheet_xml, "worksheet", "dataValidations", "dataValidation")
    wb$worksheets[[i]]$extLst <- xml_node(worksheet_xml, "worksheet", "extLst", "ext")
    wb$worksheets[[i]]$mergeCells <- xml_node(worksheet_xml, "worksheet", "mergeCells", "mergeCell")

    # wb$worksheets[[i]]$drawing <- xml_node(worksheet_xml, "worksheet", "drawing")
    wb$worksheets[[i]]$hyperlinks <- xml_node(worksheet_xml, "worksheet", "hyperlinks", "hyperlink")
    wb$worksheets[[i]]$tableParts <- xml_node(worksheet_xml, "worksheet", "tableParts", "tablePart")

    # load the data. This function reads sheet_data and returns cc and row_attr
    loadvals(wb$worksheets[[i]]$sheet_data, worksheet_xml)

  }

  ## Fix headers/footers
  for (i in seq_len(nSheets)) {
    if (!is_chart_sheet[i]) {
      if (length(wb$worksheets[[i]]$headerFooter)) {
        wb$worksheets[[i]]$headerFooter <- lapply(wb$worksheets[[i]]$headerFooter, splitHeaderFooter)
      }
    }
  }


  ##* ----------------------------------------------------------------------------------------------*##
  ### READING IN WORKSHEET DATA COMPLETE
  ##* ----------------------------------------------------------------------------------------------*##


  ## Next sheetRels to see which drawings_rels belongs to which sheet
  if (length(sheetRelsXML)) {

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
        if (is_chart_sheet[i]) {
          ind <- which(chartSheetRIds == sheetrId[i])
          rels_file <- file.path(chartSheetsRelsDir, paste0(chartsheet_rId_mapping[ind], ".rels"))
        } else {
          ind <- sheetrId[i]
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
        xml <- read_xml(allRels[[i]])
        xml <- xml_node(xml, "Relationships", "Relationship")
      } else {
        xml <- character()
      }
      return(xml)
    })

    wb$worksheets_rels <- xml





    ## Slicers -------------------------------------------------------------------------------------



    if (length(slicerXML)) {
      slicerXML <- slicerXML[order(nchar(slicerXML), slicerXML)]
      slicersFiles <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=slicer)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      inds <- lengths(slicersFiles)


      ## worksheet_rels Id for slicer will be rId0
      k <- 1L
      wb$slicers <- rep("", nSheets)
      for (i in seq_len(nSheets)) {

        ## read in slicer[j].XML sheets into sheet[i]
        if (inds[i]) {
          wb$slicers[[i]] <- slicerXML[k]
          k <- k + 1L

          # wb$worksheets_rels[[i]] <- unlist(c(
          #   wb$worksheets_rels[[i]],
          #   sprintf('<Relationship Id="rId0" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer%s.xml"/>', i)
          # ))
          wb$Content_Types <- c(
            wb$Content_Types,
            sprintf('<Override PartName="/xl/slicers/slicer%s.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>', i)
          )


          # # not sure if I want this. At least we do not create slicers?
          # slicer_xml_exists <- FALSE
          # ## Append slicer to worksheet extLst

          # if (length(wb$worksheets[[i]]$extLst)) {
          #   if (grepl('x14:slicer r:id="rId[0-9]+"', wb$worksheets[[i]]$extLst)) {
          #     wb$worksheets[[i]]$extLst <- sub('x14:slicer r:id="rId[0-9]+"', 'x14:slicer r:id="rId0"', wb$worksheets[[i]]$extLst)
          #     slicer_xml_exists <- TRUE
          #   }
          # }

          # if (!slicer_xml_exists) {
          #   wb$worksheets[[i]]$extLst <- c(wb$worksheets[[i]]$extLst, genBaseSlicerXML())
          # }
        }
      }
    }


    if (length(slicerCachesXML)) {

      ## ---- slicerCaches
      inds <- seq_along(slicerCachesXML)
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/slicerCaches/slicerCache%s.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>', inds))
      wb$slicerCaches <- sapply(slicerCachesXML[order(nchar(slicerCachesXML), slicerCachesXML)], read_xml, pointer = FALSE)
      wb$workbook.xml.rels <- c(wb$workbook.xml.rels, sprintf('<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache%s.xml"/>', 1E5 + inds, inds))
      wb$workbook$extLst <- c(wb$workbook$extLst, genSlicerCachesExtLst(1E5 + inds))
    }


    ## Tables --------------------------------------------------------------------------------------



    if (length(tablesXML)) {
      tables <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=table)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      tableSheets <- unlist(lapply(seq_along(sheetrId), function(i) rep(i, length(tables[[i]]))))

      if (length(unlist(tables))) {
        ## get the tables that belong to each worksheet and create a worksheets_rels for each
        tCount <- 2L ## table r:Ids start at 3
        for (i in seq_along(tables)) {
          if (length(tables[[i]])) {
            k <- seq_along(tables[[i]]) + tCount
            # wb$worksheets_rels[[i]] <- unlist(c(
            #   wb$worksheets_rels[[i]],
            #   sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>', k, k)
            # ))


            #wb$worksheets[[i]]$tableParts <- sprintf("<tablePart r:id=\"rId%s\"/>", k)
            tCount <- tCount + length(k)
          }
        }

        ## sort the tables into the order they appear in the xml and tables variables
        names(tablesXML) <- basename(tablesXML)
        tablesXML <- tablesXML[sprintf("table%s.xml", unlist(tables))]

        ## tables are now in correct order so we can read them in as they are
        wb$tables <- sapply(tablesXML, read_xml, pointer = FALSE)

        ## pull out refs and attach names
        refs <- regmatches(wb$tables, regexpr('(?<=ref=")[0-9A-Z:]+', wb$tables, perl = TRUE))
        names(wb$tables) <- refs

        wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', seq_along(wb$tables)))

        ## relabel ids
        for (i in seq_along(wb$tables)) {
          newId <- sprintf(' id="%s" ', i + 2)
          wb$tables[[i]] <- sub(' id="[0-9]+" ', newId, wb$tables[[i]])
        }

        displayNames <- unlist(regmatches(wb$tables, regexpr('(?<=displayName=").*?[^"]+', wb$tables, perl = TRUE)))
        if (length(displayNames) != length(tablesXML)) {
          displayNames <- paste0("Table", seq_along(tablesXML))
        }

        attr(wb$tables, "sheet") <- tableSheets
        attr(wb$tables, "tableName") <- displayNames

        for (i in seq_along(tableSheets)) {
          table_sheet_i <- tableSheets[i]
          attr(wb$worksheets[[table_sheet_i]]$tableParts, "tableName") <- c(attr(wb$worksheets[[table_sheet_i]]$tableParts, "tableName"), displayNames[i])
        }
      }
    } ## if(length(tablesXML))

    ## might we have some external hyperlinks
    # TODO use lengths()
    if (any(sapply(wb$worksheets[!is_chart_sheet], function(x) length(x$hyperlinks)) > 0)) {

      ## Do we have external hyperlinks
      hlinks <- lapply(xml, function(x) x[grepl("hyperlink", x) & grepl("External", x)])
      # TODO use lengths()
      hlinksInds <- which(lengths(hlinks) > 0)

      ## If it's an external hyperlink it will have a target in the sheet_rels
      if (length(hlinksInds)) {
        for (i in hlinksInds) {
          ids <- unlist(lapply(hlinks[[i]], function(x) regmatches(x, gregexpr('(?<=Id=").*?"', x, perl = TRUE))[[1]]))
          ids <- gsub('"$', "", ids)

          targets <- unlist(lapply(hlinks[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          targets <- gsub('"$', "", targets)

          ids2 <- lapply(wb$worksheets[[i]]$hyperlinks, function(x) regmatches(x, gregexpr('(?<=r:id=").*?"', x, perl = TRUE))[[1]])
          ids2[lengths(ids2) == 0] <- NA
          ids2 <- gsub('"$', "", unlist(ids2))

          targets <- targets[match(ids2, ids)]
          names(wb$worksheets[[i]]$hyperlinks) <- targets
        }
      }
    }



    ## Drawings ------------------------------------------------------------------------------------



    ## xml is in the order of the sheets, drawIngs is toes to sheet position of hasDrawing
    ## Not every sheet has a drawing.xml

    drawXMLrelationship <- lapply(xml, function(x) grep("drawings/drawing", x, value = TRUE))
    # TODO use lengths()
    hasDrawing <- lengths(drawXMLrelationship) > 0 ## which sheets have a drawing

    if (length(drawingRelsXML)) {
      dRels <- lapply(drawingRelsXML, read_xml, pointer = FALSE)
      # TODO lapply xml_node Relationships?
      dRels <- gsub("<Relationships .*?>", "", dRels)
      dRels <- gsub("</Relationships>", "", dRels)
    }

    if (length(drawingsXML)) {
      dXML <- lapply(drawingsXML, read_xml, pointer = FALSE)
      # this creates crippled drawings files
      dXML <- gsub("<xdr:wsDr .*?>", "", dXML)
      dXML <- gsub("</xdr:wsDr>", "", dXML)


      for (drawing in seq_along(drawingsXML)) {
        drwng_xml <- read_xml(drawingsXML[drawing])
        wb$drawings[[drawing]] <- xml_node(drwng_xml, "xdr:wsDr")
      }

      # ptn1 <- "<(mc:AlternateContent|xdr:oneCellAnchor|xdr:twoCellAnchor|xdr:absoluteAnchor)"
      # ptn2 <- "</(mc:AlternateContent|xdr:oneCellAnchor|xdr:twoCellAnchor|xdr:absoluteAnchor)>"

      # ## split at one/two cell Anchor
      # dXML <- regmatches(dXML, gregexpr(paste0(ptn1, ".*?", ptn2), dXML))
    }


    # loop over all worksheets and assign drawing to sheet
    if (any(hasDrawing)) {
      for (i in seq_along(xml)) {
        if (hasDrawing[i]) {
          target <- unlist(lapply(drawXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          target <- basename(gsub('"$', "", target))

          ## sheet_i has which(hasDrawing)[[i]]
          relsInd <- grepl(target, drawingRelsXML)
          if (any(relsInd)) {
            wb$drawings_rels[i] <- dRels[relsInd]
          }

          drawingInd <- grepl(target, drawingsXML)
          if (any(drawingInd)) {
            wb$drawings[i] <- sprintf("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">%s</xdr:wsDr>", dXML[drawingInd])
          }
        }
      }
    }




    ## VML Drawings --------------------------------------------------------------------------------


    if (length(vmlDrawingXML)) {
      wb$Content_Types <- c(wb$Content_Types, '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>')

      # TODO missed <<-
      drawXMLrelationship <<- lapply(xml, function(x) grep("drawings/vmlDrawing", x, value = TRUE))

      for (i in seq_along(vmlDrawingXML)) {
        wb$drawings_vml[[i]] <- read_xml(vmlDrawingXML[[i]], pointer = FALSE)
      }


      # TODO use lengths()
      hasDrawing <- lengths(drawXMLrelationship) > 0 ## which sheets have a drawing

      ## loop over all worksheets and assign drawing to sheet
      if (any(hasDrawing)) {
        for (i in seq_along(xml)) {
          if (hasDrawing[i]) {
            target <- unlist(lapply(drawXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
            target <- basename(gsub('"$', "", target))
            ind <- grepl(target, vmlDrawingXML)

            if (any(ind)) {
              txt <- read_xml(vmlDrawingXML[ind], pointer = FALSE)

              i1 <- regexpr("<v:shapetype", txt, fixed = TRUE)
              i2 <- regexpr("</xml>", txt, fixed = TRUE)

              wb$vml[[i]] <- substring(text = txt, first = i1, last = (i2 - 1L))

              relsInd <- grepl(target, vmlDrawingRelsXML)
              if (any(relsInd)) {
                wb$vml_rels[i] <- vmlDrawingRelsXML[relsInd]
              }
            }
          }
        }
      }
    }







    ## vmlDrawing and comments
    if (length(commentsXML)) {

      com_rId <- vector("list", length(commentsrelXML))
      names(com_rId) <- commentsrelXML
      for (com_rel in commentsrelXML) {
        rel_xml <- read_xml(com_rel)
        attrs <- xml_attribute(rel_xml, "Relationships", "Relationship")
        rel <- rbindlist(attrs)
        com_rId[[com_rel]] <- rel
      }

      drawXMLrelationship <- lapply(xml, function(x) grep("drawings/vmlDrawing[0-9]+\\.vml", x, value = TRUE))
      hasDrawing <- lengths(drawXMLrelationship) > 0 ## which sheets have a drawing

      commentXMLrelationship <- lapply(xml, function(x) grep("comments[0-9]+\\.xml", x, value = TRUE))
      hasComment <- lengths(commentXMLrelationship) > 0 ## which sheets have a comment

      for (i in seq_along(xml)) {
        if (hasComment[i]) {
          target <- unlist(lapply(drawXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          target <- basename(gsub('"$', "", target))
          ind <- grepl(target, vmlDrawingXML)

          if (any(ind)) {
            txt <- read_xml(vmlDrawingXML[ind], pointer = FALSE)

            cd <- unique(xml_node(txt, "xml", "*", "x:ClientData"))
            cd <- grep('ObjectType="Note"', cd, value = TRUE)
            cd <- paste0(cd, ">")

            ## now loada comment
            target <- unlist(lapply(commentXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
            target <- basename(gsub('"$', "", target))

            txt <- read_xml(grep(target, commentsXML, value = TRUE))


            authors <- xml_node(txt, "comments", "authors", "author")
            authors <- gsub("<author>|</author>", "", authors)

            comments <- xml_node(txt, "comments", "commentList", "comment")

            refs <- regmatches(comments, regexpr('(?<=ref=").*?[^"]+', comments, perl = TRUE))

            authorsInds <- as.integer(regmatches(comments, regexpr('(?<=authorId=").*?[^"]+', comments, perl = TRUE))) + 1
            authors <- authors[authorsInds]

            style <- lapply(comments, function(x) unlist(xml_node(x, "comment", "text", "r", "rPr")) )

            comments <- regmatches(comments,
              gregexpr("(?<=<t( |>))[\\s\\S]+?(?=</t>)", comments, perl = TRUE))
            comments <- lapply(comments, function(x) gsub(".*?>", "", x, perl = TRUE))


            wb$comments[[i]] <- lapply(seq_along(comments), function(j) {
              comment_list <- list(
                #"refId" = com_rId[j],
                "ref" = refs[j],
                "author" = authors[j],
                "comment" = comments[[j]],
                "style" = style[[j]],
                "clientData" = cd[[j]]
              )
            })
          }
        }
      }
    }

    ## Threaded comments
    if (length(threadCommentsXML) > 0) {
      threadCommentsXMLrelationship <- lapply(xml, function(x) grep("threadedComment[0-9]+\\.xml", x, value = TRUE))
      hasThreadComments <- lengths(threadCommentsXMLrelationship) > 0
      if(any(hasThreadComments)) {
        for (i in seq_along(xml)) {
          if (hasThreadComments[i]) {
            target <- unlist(lapply(threadCommentsXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
            target <- basename(gsub('"$', "", target))

            wb$threadComments[[i]] <- grep(target, threadCommentsXML, value = TRUE)

          }
        }
      }
      wb$Content_Types <- c(
        wb$Content_Types,
        sprintf('<Override PartName="/xl/threadedComments/%s" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>',
          sapply(threadCommentsXML, basename))
      )
    }

    ## Persons (needed for Threaded Comment)
    if(length(personXML) > 0){
      wb$persons <- personXML
      wb$Content_Types <- c(
        wb$Content_Types,
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>'
      )
      wb$workbook.xml.rels <- c(
        wb$workbook.xml.rels,
        '<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>')
    }


    ## rels image
    drawXMLrelationship <- lapply(xml, function(x) grep("relationships/image", x, value = TRUE))
    hasDrawing <- lengths(drawXMLrelationship) > 0 ## which sheets have a drawing
    if (any(hasDrawing)) {
      for (i in seq_along(xml)) {
        if (hasDrawing[i]) {
          image_ids <- unlist(getId(drawXMLrelationship[[i]]))
          new_image_ids <- paste0("rId", seq_along(image_ids) + 70000)
          for (j in seq_along(image_ids)) {
            wb$worksheets[[i]]$oleObjects <- gsub(image_ids[j], new_image_ids[j], wb$worksheets[[i]]$oleObjects, fixed = TRUE)
            # wb$worksheets_rels[[i]] <- c(wb$worksheets_rels[[i]], gsub(image_ids[j], new_image_ids[j], drawXMLrelationship[[i]][j], fixed = TRUE))
          }
        }
      }
    }

    ## rels image
    drawXMLrelationship <- lapply(xml, function(x) grep("relationships/package", x, value = TRUE))
    hasDrawing <- lengths(drawXMLrelationship) > 0 ## which sheets have a drawing
    if (any(hasDrawing)) {
      for (i in seq_along(xml)) {
        if (hasDrawing[i]) {
          image_ids <- unlist(getId(drawXMLrelationship[[i]]))
          new_image_ids <- paste0("rId", seq_along(image_ids) + 90000)
          for (j in seq_along(image_ids)) {
            wb$worksheets[[i]]$oleObjects <- gsub(image_ids[j], new_image_ids[j], wb$worksheets[[i]]$oleObjects, fixed = TRUE)
            # wb$worksheets_rels[[i]] <- c(
            #   wb$worksheets_rels[[i]],
            #   sprintf("<Relationship Id=\"%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package\" Target=\"../embeddings/Microsoft_Word_Document1.docx\"/>", new_image_ids[j])
            # )
          }
        }
      }
    }



    ## Embedded docx
    if (length(embeddings) > 0) {
      wb$Content_Types <- c(wb$Content_Types, '<Default Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"/>')
      wb$embeddings <- embeddings
    }



    # ## pivot tables
    # if (length(pivotTableXML) > 0) {
    #   pivotTableJ <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=pivotTable)[0-9]+(?=\\.xml)", x, perl = TRUE))))
    #   sheetWithPivot <- which(lengths(pivotTableJ) > 0)

    #   pivotRels <- lapply(xml, function(x) {
    #     y <- grep("pivotTable", x, value = TRUE)
    #     y[order(nchar(y), y)]
    #   })
    #   hasPivot <- lengths(pivotRels) > 0

    #   ## Modify rIds
    #   for (i in seq_along(pivotRels)) {
    #     if (hasPivot[i]) {
    #       for (j in seq_along(pivotRels[[i]])) {
    #         pivotRels[[i]][j] <- gsub('"rId[0-9]+"', sprintf('"rId%s"', 20000L + j), pivotRels[[i]][j])
    #       }

    #       wb$worksheets_rels[[i]] <- c(wb$worksheets_rels[[i]], pivotRels[[i]])
    #     }
    #   }


    #   ## remove any workbook_res references to pivot tables that are not being used in worksheet_rels
    #   inds <- seq_along(wb$pivotTables.xml.rels)
    #   fileNo <- as.integer(unlist(regmatches(unlist(wb$worksheets_rels), gregexpr("(?<=pivotTable)[0-9]+(?=\\.xml)", unlist(wb$worksheets_rels), perl = TRUE))))
    #   inds <- inds[!inds %in% fileNo]

    #   if (length(inds) > 0) {
    #     toRemove <- paste(sprintf("(pivotCacheDefinition%s\\.xml)", inds), collapse = "|")
    #     fileNo <- grep(toRemove, wb$pivotTables.xml.rels)
    #     toRemove <- paste(sprintf("(pivotCacheDefinition%s\\.xml)", fileNo), collapse = "|")

    #     ## remove reference to file from workbook.xml.res
    #     wb$workbook.xml.rels <- wb$workbook.xml.rels[!grepl(toRemove, wb$workbook.xml.rels)]
    #   }
    # }
  } ## end of worksheetRels

  ## convert hyperliks to hyperlink objects
  for (i in seq_len(nSheets)) {
    wb$worksheets[[i]]$hyperlinks <- xml_to_hyperlink(wb$worksheets[[i]]$hyperlinks)
  }



  ## queryTables
  if (length(queryTablesXML) > 0) {
    ids <- as.numeric(regmatches(queryTablesXML, regexpr("[0-9]+(?=\\.xml)", queryTablesXML, perl = TRUE)))
    wb$queryTables <- unlist(lapply(queryTablesXML[order(ids)], read_xml, pointer = FALSE))
    wb$Content_Types <- c(
      wb$Content_Types,
      sprintf('<Override PartName="/xl/queryTables/queryTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>', seq_along(queryTablesXML))
    )
  }


  ## connections
  if (length(connectionsXML) > 0) {
    wb$connections <- read_xml(connectionsXML, pointer = FALSE)
    wb$workbook.xml.rels <- c(wb$workbook.xml.rels, '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>')
    wb$Content_Types <- c(wb$Content_Types, '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>')
  }




  ## table rels
  if (length(tableRelsXML) > 0) {

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
