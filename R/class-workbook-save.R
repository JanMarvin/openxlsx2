workbook_save <- function(self, private, path = self$path, overwrite = TRUE) {

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
}
