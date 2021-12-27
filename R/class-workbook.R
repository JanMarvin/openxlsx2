
Workbook <- setRefClass(
  "Workbook",
  fields = c(
    "sheet_names" = "character",

    "charts" = "ANY",
    "isChartSheet" = "logical",

    "colOutlineLevels" = "ANY",
    "colWidths" = "ANY",
    "connections" = "ANY",
    "Content_Types" = "character",
    "core" = "character",
    "drawings" = "ANY",
    "drawings_rels" = "ANY",
    "drawings_vml" = "ANY",
    "embeddings" = "ANY",
    "externalLinks" = "ANY",
    "externalLinksRels" = "ANY",

    "headFoot" = "ANY",
    "media" = "ANY",
    "outlineLevels" = "ANY",

    "persons" = "ANY",

    "pivotTables" = "ANY",
    "pivotTables.xml.rels" = "ANY",
    "pivotDefinitions" = "ANY",
    "pivotRecords" = "ANY",
    "pivotDefinitionsRels" = "ANY",

    "queryTables" = "ANY",
    "rowHeights" = "ANY",

    "slicers" = "ANY",
    "slicerCaches" = "ANY",

    "sharedStrings" = "ANY",
    "styleObjects" = "ANY",

    "styles" = "ANY",
    "styles_xml" = "ANY",
    "tables" = "ANY",
    "tables.xml.rels" = "ANY",
    "theme" = "ANY",

    "vbaProject" = "ANY",
    "vml" = "ANY",
    "vml_rels" = "ANY",
    "comments" = "ANY",
    "threadComments" = "ANY",

    "workbook" = "ANY",
    "workbook.xml.rels" = "ANY",
    "worksheets" = "ANY",
    "worksheets_rels" = "ANY",
    "sheetOrder" = "integer"
  ),

  methods = list(

    initialize = function(
      creator = "",
      title = NULL,
      subject = NULL,
      category = NULL
    ) {
      .self$charts <- list()
      .self$isChartSheet <- logical(0)

      .self$colWidths <- list()
      .self$colOutlineLevels <- list()
      attr(.self$colOutlineLevels, "hidden") <- NULL
      .self$connections <- NULL
      .self$Content_Types <- genBaseContent_Type()
      .self$core <-
        genBaseCore(
          creator = creator,
          title = title,
          subject = subject,
          category = category
        )
      .self$comments <- list()
      .self$threadComments <- list()


      .self$drawings <- list()
      .self$drawings_rels <- list()
      .self$drawings_vml <- list()

      .self$embeddings <- NULL
      .self$externalLinks <- NULL
      .self$externalLinksRels <- NULL

      .self$headFoot <- NULL

      .self$media <- list()

      .self$persons <- NULL

      .self$pivotTables <- NULL
      .self$pivotTables.xml.rels <- NULL
      .self$pivotDefinitions <- NULL
      .self$pivotRecords <- NULL
      .self$pivotDefinitionsRels <- NULL

      .self$queryTables <- NULL
      .self$rowHeights <- list()
      .self$outlineLevels <- list()
      attr(.self$outlineLevels, "hidden") <- NULL

      .self$slicers <- NULL
      .self$slicerCaches <- NULL

      .self$sheet_names <- character(0)
      .self$sheetOrder <- integer(0)

      .self$sharedStrings <- list()
      attr(.self$sharedStrings, "uniqueCount") <- 0

      .self$styles <- genBaseStyleSheet()
      .self$styleObjects <- list()


      .self$tables <- NULL
      .self$tables.xml.rels <- NULL
      .self$theme <- NULL


      .self$vbaProject <- NULL
      .self$vml <- list()
      .self$vml_rels <- list()



      .self$workbook <- genBaseWorkbook()
      .self$workbook.xml.rels <- genBaseWorkbook.xml.rels()

      .self$worksheets <- list()
      .self$worksheets_rels <- list()
    },


    addWorksheet = function(sheetName,
      showGridLines = TRUE,
      tabColour = NULL,
      zoom = 100,
      oddHeader = NULL,
      oddFooter = NULL,
      evenHeader = NULL,
      evenFooter = NULL,
      firstHeader = NULL,
      firstFooter = NULL,
      visible = TRUE,
      hasDrawing = FALSE,
      paperSize = 9,
      orientation = "portrait",
      hdpi = 300,
      vdpi = 300) {
      if (!missing(sheetName)) {
        if (grepl(pattern = ":", x = sheetName)) {
          stop("colon not allowed in sheet names in Excel")
        }
      }
      newSheetIndex <- length(worksheets) + 1L

      if (newSheetIndex > 1) {
        sheetId <-
          max(as.integer(regmatches(
            workbook$sheets,
            regexpr('(?<=sheetId=")[0-9]+', workbook$sheets, perl = TRUE)
          ))) + 1L
      } else {
        sheetId <- 1
      }


      ## fix visible value
      visible <- tolower(visible)
      if (visible == "true") {
        visible <- "visible"
      }

      if (visible == "false") {
        visible <- "hidden"
      }

      if (visible == "veryhidden") {
        visible <- "veryHidden"
      }


      ##  Add sheet to workbook.xml
      .self$workbook$sheets <-
        c(
          workbook$sheets,
          sprintf(
            '<sheet name="%s" sheetId="%s" state="%s" r:id="rId%s"/>',
            sheetName,
            sheetId,
            visible,
            newSheetIndex
          )
        )

      ## append to worksheets list
      .self$worksheets <-
        append(
          worksheets,
          Worksheet$new(
            showGridLines = showGridLines,
            tabSelected = newSheetIndex == 1,
            tabColour = tabColour,
            zoom = zoom,
            oddHeader = oddHeader,
            oddFooter = oddFooter,
            evenHeader = evenHeader,
            evenFooter = evenFooter,
            firstHeader = firstHeader,
            firstFooter = firstFooter,
            paperSize = paperSize,
            orientation = orientation,
            hdpi = hdpi,
            vdpi = vdpi
          )
        )


      ## update content_tyes
      ## add a drawing.xml for the worksheet
      if (hasDrawing) {
        .self$Content_Types <-
          c(
            Content_Types,
            sprintf(
              '<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
              newSheetIndex
            ),
            sprintf(
              '<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
              newSheetIndex
            )
          )
      } else {
        .self$Content_Types <-
          c(
            Content_Types,
            sprintf(
              '<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
              newSheetIndex
            )
          )
      }

      ## Update xl/rels
      .self$workbook.xml.rels <- c(
        workbook.xml.rels,
        sprintf(
          '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>',
          newSheetIndex
        )
      )


      ## create sheet.rels to simplify id assignment
      .self$worksheets_rels[[newSheetIndex]] <-
        genBaseSheetRels(newSheetIndex)
      .self$drawings_rels[[newSheetIndex]] <- list()
      .self$drawings[[newSheetIndex]] <- list()

      .self$vml_rels[[newSheetIndex]] <- list()
      .self$vml[[newSheetIndex]] <- list()

      .self$isChartSheet[[newSheetIndex]] <- FALSE
      .self$comments[[newSheetIndex]] <- list()
      .self$threadComments[[newSheetIndex]] <- list()

      .self$rowHeights[[newSheetIndex]] <- list()
      .self$colWidths[[newSheetIndex]] <- list()
      .self$colOutlineLevels[[newSheetIndex]] <- list()
      .self$outlineLevels[[newSheetIndex]] <- list()

      .self$sheetOrder <- c(sheetOrder, as.integer(newSheetIndex))
      .self$sheet_names <- c(sheet_names, sheetName)

      invisible(newSheetIndex)
    },

    cloneWorksheet = function(sheetName, clonedSheet) {
      clonedSheet <- validateSheet(clonedSheet)
      if (!missing(sheetName)) {
        if (grepl(pattern = ":", x = sheetName)) {
          stop("colon not allowed in sheet names in Excel")
        }
      }
      newSheetIndex <- length(worksheets) + 1L
      if (newSheetIndex > 1) {
        sheetId <-
          max(as.integer(regmatches(
            workbook$sheets,
            regexpr('(?<=sheetId=")[0-9]+', workbook$sheets, perl = TRUE)
          ))) + 1L
      } else {
        sheetId <- 1
      }


      ## copy visibility from cloned sheet!
      visible <-
        regmatches(
          workbook$sheets[[clonedSheet]],
          regexpr('(?<=state=")[^"]+', workbook$sheets[[clonedSheet]], perl = TRUE)
        )

      ##  Add sheet to workbook.xml
      .self$workbook$sheets <-
        c(
          workbook$sheets,
          sprintf(
            '<sheet name="%s" sheetId="%s" state="%s" r:id="rId%s"/>',
            sheetName,
            sheetId,
            visible,
            newSheetIndex
          )
        )

      ## append to worksheets list
      .self$worksheets <-
        append(worksheets, worksheets[[clonedSheet]]$copy())


      ## update content_tyes
      ## add a drawing.xml for the worksheet
      .self$Content_Types <-
        c(
          Content_Types,
          sprintf(
            '<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            newSheetIndex
          ),
          sprintf(
            '<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
            newSheetIndex
          )
        )

      ## Update xl/rels
      .self$workbook.xml.rels <- c(
        workbook.xml.rels,
        sprintf(
          '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>',
          newSheetIndex
        )
      )

      ## create sheet.rels to simplify id assignment
      .self$worksheets_rels[[newSheetIndex]] <-
        genBaseSheetRels(newSheetIndex)
      .self$drawings_rels[[newSheetIndex]] <- drawings_rels[[clonedSheet]]

      # give each chart its own filename (images can re-use the same file, but charts can't)
      .self$drawings_rels[[newSheetIndex]] <-
        sapply(drawings_rels[[newSheetIndex]], function(rl) {
          chartfiles <-
            regmatches(
              rl,
              gregexpr("(?<=charts/)chart[0-9]+\\.xml", rl, perl = TRUE)
            )[[1]]
          for (cf in chartfiles) {
            chartid <- length(charts) + 1
            newname <- stri_join("chart", chartid, ".xml")
            fl <- charts[cf]

            # Read the chartfile and adjust all formulas to point to the new
            # sheet name instead of the clone source
            # The result is saved to a new chart xml file
            newfl <- file.path(dirname(fl), newname)
            .self$charts[newname] <- newfl
            chart <- readUTF8(fl)
            chart <-
              gsub(
                stri_join("(?<=')", sheet_names[[clonedSheet]], "(?='!)"),
                stri_join("'", sheetName, "'"),
                chart,
                perl = TRUE
              )
            chart <-
              gsub(
                stri_join("(?<=[^A-Za-z0-9])", sheet_names[[clonedSheet]], "(?=!)"),
                stri_join("'", sheetName, "'"),
                chart,
                perl = TRUE
              )
            writeLines(chart, newfl)
            # file.copy(fl, newfl)
            .self$Content_Types <-
              c(
                Content_Types,
                sprintf(
                  '<Override PartName="/xl/charts/%s" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
                  newname
                )
              )
            rl <- gsub(stri_join("(?<=charts/)", cf), newname, rl, perl = TRUE)
          }
          rl
        }, USE.NAMES = FALSE)
      # The IDs in the drawings array are sheet-specific, so within the new cloned sheet
      # the same IDs can be used => no need to modify drawings
      .self$drawings[[newSheetIndex]] <- drawings[[clonedSheet]]

      .self$vml_rels[[newSheetIndex]] <- vml_rels[[clonedSheet]]
      .self$vml[[newSheetIndex]] <- vml[[clonedSheet]]

      .self$isChartSheet[[newSheetIndex]] <- isChartSheet[[clonedSheet]]
      .self$comments[[newSheetIndex]] <- comments[[clonedSheet]]
      .self$threadComments[[newSheetIndex]] <- threadComments[[clonedSheet]]

      .self$rowHeights[[newSheetIndex]] <- rowHeights[[clonedSheet]]
      .self$colWidths[[newSheetIndex]] <- colWidths[[clonedSheet]]

      .self$colOutlineLevels[[newSheetIndex]] <- colOutlineLevels[[clonedSheet]]
      .self$outlineLevels[[newSheetIndex]] <- outlineLevels[[clonedSheet]]

      .self$sheetOrder <- c(sheetOrder, as.integer(newSheetIndex))
      .self$sheet_names <- c(sheet_names, sheetName)


      ############################
      ## STYLE
      ## ... objects are stored in a global list, so we need to get all styles
      ## assigned to the cloned sheet and duplicate them
      sheetStyles <- Filter(function(s) {
        s$sheet == sheet_names[[clonedSheet]]
      }, styleObjects)
      .self$styleObjects <- c(
        styleObjects,
        Map(function(s) {
          s$sheet <- sheetName
          s
        }, sheetStyles)
      )


      ############################
      ## TABLES
      ## ... are stored in the $tables list, with the name and sheet as attr
      ## and in the worksheets[]$tableParts list. We also need to adjust the
      ## worksheets_rels and set the content type for the new table

      tbls <- tables[attr(tables, "sheet") == clonedSheet]
      for (t in tbls) {
        # Extract table name, displayName and ID from the xml
        oldname <- regmatches(t, regexpr('(?<= name=")[^"]+', t, perl = TRUE))
        olddispname <- regmatches(t, regexpr('(?<= displayName=")[^"]+', t, perl = TRUE))
        oldid <- regmatches(t, regexpr('(?<= id=")[^"]+', t, perl = TRUE))
        ref <- regmatches(t, regexpr('(?<= ref=")[^"]+', t, perl = TRUE))

        # Find new, unused table names by appending _n, where n=1,2,...
        n <- 0
        while (stri_join(oldname, "_", n) %in% attr(tables, "tableName")) {
          n <- n + 1
        }
        newname <- stri_join(oldname, "_", n)
        newdispname <- stri_join(olddispname, "_", n)
        newid <- as.character(length(tables) + 3L)

        # Use the table definition from the cloned sheet and simply replace the names
        newt <- t
        newt <-
          gsub(
            stri_join(" name=\"", oldname, "\""),
            stri_join(" name=\"", newname, "\""),
            newt
          )
        newt <-
          gsub(
            stri_join(" displayName=\"", olddispname, "\""),
            stri_join(" displayName=\"", newdispname, "\""),
            newt
          )
        newt <-
          gsub(
            stri_join("(<table [^<]* id=\")", oldid, "\""),
            stri_join("\\1", newid, "\""),
            newt
          )

        oldtables <- tables
        .self$tables <- c(oldtables, newt)
        names(.self$tables) <- c(names(oldtables), ref)
        attr(.self$tables, "sheet") <-
          c(attr(oldtables, "sheet"), newSheetIndex)
        attr(.self$tables, "tableName") <-
          c(attr(oldtables, "tableName"), newname)

        oldparts <- worksheets[[newSheetIndex]]$tableParts
        .self$worksheets[[newSheetIndex]]$tableParts <-
          c(oldparts, sprintf('<tablePart r:id="rId%s"/>', newid))
        attr(.self$worksheets[[newSheetIndex]]$tableParts, "tableName") <-
          c(attr(oldparts, "tableName"), newname)
        names(attr(.self$worksheets[[newSheetIndex]]$tableParts, "tableName")) <-
          c(names(attr(oldparts, "tableName")), ref)

        .self$Content_Types <-
          c(
            Content_Types,
            sprintf(
              '<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
              newid
            )
          )
        .self$tables.xml.rels <- append(tables.xml.rels, "")

        .self$worksheets_rels[[newSheetIndex]] <-
          c(
            worksheets_rels[[newSheetIndex]],
            sprintf(
              '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',
              newid,
              newid
            )
          )
      }

      # TODO: The following items are currently NOT copied/duplicated for the cloned sheet:
      #   - Comments
      #   - Pivot tables

      invisible(newSheetIndex)
    },

    addChartSheet = function(sheetName,
      tabColour = NULL,
      zoom = 100) {
      newSheetIndex <- length(worksheets) + 1L

      if (newSheetIndex > 1) {
        sheetId <-
          max(as.integer(regmatches(
            workbook$sheets,
            regexpr('(?<=sheetId=")[0-9]+', workbook$sheets, perl = TRUE)
          ))) + 1L
      } else {
        sheetId <- 1
      }

      ##  Add sheet to workbook.xml
      .self$workbook$sheets <-
        c(
          workbook$sheets,
          sprintf(
            '<sheet name="%s" sheetId="%s" r:id="rId%s"/>',
            sheetName,
            sheetId,
            newSheetIndex
          )
        )

      ## append to worksheets list
      .self$worksheets <-
        append(
          worksheets,
          ChartSheet$new(
            tabSelected = newSheetIndex == 1,
            tabColour = tabColour,
            zoom = zoom
          )
        )
      .self$sheet_names <- c(sheet_names, sheetName)

      ## update content_tyes
      .self$Content_Types <-
        c(
          Content_Types,
          sprintf(
            '<Override PartName="/xl/chartsheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>',
            newSheetIndex
          )
        )

      ## Update xl/rels
      .self$workbook.xml.rels <- c(
        workbook.xml.rels,
        sprintf(
          '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet%s.xml"/>',
          newSheetIndex
        )
      )



      ## add a drawing.xml for the worksheet
      .self$Content_Types <-
        c(
          Content_Types,
          sprintf(
            '<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
            newSheetIndex
          )
        )

      ## create sheet.rels to simplify id assignment
      .self$worksheets_rels[[newSheetIndex]] <-
        genBaseSheetRels(newSheetIndex)
      .self$drawings_rels[[newSheetIndex]] <- list()
      .self$drawings[[newSheetIndex]] <- list()

      .self$isChartSheet[[newSheetIndex]] <- TRUE

      .self$rowHeights[[newSheetIndex]] <- list()
      .self$colWidths[[newSheetIndex]] <- list()

      .self$colOutlineLevels[[newSheetIndex]] <- list()
      .self$outlineLevels[[newSheetIndex]] <- list()

      .self$vml_rels[[newSheetIndex]] <- list()
      .self$vml[[newSheetIndex]] <- list()

      .self$sheetOrder <- c(sheetOrder, newSheetIndex)

      invisible(newSheetIndex)
    },

    saveWorkbook = function() {
      ## temp directory to save XML files prior to compressing
      tmpDir <- file.path(tempfile(pattern = "workbookTemp_"))

      if (file.exists(tmpDir)) {
        unlink(tmpDir, recursive = TRUE, force = TRUE)
      }

      success <- dir.create(path = tmpDir, recursive = TRUE)
      if (!success) {
        stop(sprintf("Failed to create temporary directory '%s'", tmpDir))
      }

      .self$preSaveCleanUp()

      nSheets <- length(worksheets)
      nThemes <- length(theme)
      nPivots <- length(pivotDefinitions)
      nSlicers <- length(slicers)
      nComments <- sum(sapply(comments, length) > 0)
      nThreadComments <- sum(sapply(threadComments, length) > 0)
      nPersons <- length(persons)
      nVML <- sum(sapply(vml, length) > 0)

      relsDir <- file.path(tmpDir, "_rels")
      dir.create(path = relsDir, recursive = TRUE)

      docPropsDir <- file.path(tmpDir, "docProps")
      dir.create(path = docPropsDir, recursive = TRUE)

      xlDir <- file.path(tmpDir, "xl")
      dir.create(path = xlDir, recursive = TRUE)

      xlrelsDir <- file.path(tmpDir, "xl", "_rels")
      dir.create(path = xlrelsDir, recursive = TRUE)

      xlTablesDir <- file.path(tmpDir, "xl", "tables")
      dir.create(path = xlTablesDir, recursive = TRUE)

      xlTablesRelsDir <- file.path(xlTablesDir, "_rels")
      dir.create(path = xlTablesRelsDir, recursive = TRUE)

      if (length(media) > 0) {
        xlmediaDir <- file.path(tmpDir, "xl", "media")
        dir.create(path = xlmediaDir, recursive = TRUE)
      }


      ## will always have a theme
      xlthemeDir <- file.path(tmpDir, "xl", "theme")
      dir.create(path = xlthemeDir, recursive = TRUE)

      if (is.null(theme)) {
        con <- file(file.path(xlthemeDir, "theme1.xml"), open = "wb")
        writeBin(charToRaw(genBaseTheme()), con)
        close(con)
      } else {
        lapply(1:nThemes, function(i) {
          con <-
            file(file.path(xlthemeDir, stri_join("theme", i, ".xml")), open = "wb")
          writeBin(charToRaw(pxml(theme[[i]])), con)
          close(con)
        })
      }


      ## will always have drawings
      xlworksheetsDir <- file.path(tmpDir, "xl", "worksheets")
      dir.create(path = xlworksheetsDir, recursive = TRUE)

      xlworksheetsRelsDir <-
        file.path(tmpDir, "xl", "worksheets", "_rels")
      dir.create(path = xlworksheetsRelsDir, recursive = TRUE)

      xldrawingsDir <- file.path(tmpDir, "xl", "drawings")
      dir.create(path = xldrawingsDir, recursive = TRUE)

      xldrawingsRelsDir <- file.path(tmpDir, "xl", "drawings", "_rels")
      dir.create(path = xldrawingsRelsDir, recursive = TRUE)

      ## charts
      if (length(charts) > 0) {
        file.copy(
          from = dirname(charts[1]),
          to = file.path(tmpDir, "xl"),
          recursive = TRUE
        )
      }


      ## xl/comments.xml
      if (nComments > 0 | nVML > 0) {
        for (i in 1:nSheets) {
          if (length(comments[[i]]) > 0) {
            fn <- sprintf("comments%s.xml", i)

            .self$Content_Types <- c(
              Content_Types,
              sprintf(
                '<Override PartName="/xl/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>',
                fn
              )
            )

            .self$worksheets_rels[[i]] <- unique(c(
              worksheets_rels[[i]],
              sprintf(
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../%s"/>',
                fn
              )
            ))

            writeCommentXML(
              comment_list = comments[[i]],
              file_name = file.path(tmpDir, "xl", fn)
            )
          }
        }

        .self$writeDrawingVML(xldrawingsDir)
      }

      ## Threaded Comments xl/threadedComments/threadedComment.xml
      if (nThreadComments > 0){
        xlThreadComments <- file.path(tmpDir, "xl", "threadedComments")
        dir.create(path = xlThreadComments, recursive = TRUE)

        for (i in seq_len(nSheets)) {
          if (length(threadComments[[i]]) > 0) {
            fl <- threadComments[[i]]
            file.copy(
              from = fl,
              to = file.path(xlThreadComments, basename(fl)),
              overwrite = TRUE,
              copy.date = TRUE
            )

            .self$worksheets_rels[[i]] <- unique(c(
              worksheets_rels[[i]],
              sprintf(
                '<Relationship Id="rIdthread" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/%s"/>',
                basename(fl)
              )
            ))
          }
        }
      }

      ## xl/persons/person.xml
      if (nPersons > 0){
        personDir <- file.path(tmpDir, "xl", "persons")
        dir.create(path = personDir, recursive = TRUE)
        file.copy(
          from = persons,
          to = personDir,
          overwrite = TRUE
        )

      }



      if (length(embeddings) > 0) {
        embeddingsDir <- file.path(tmpDir, "xl", "embeddings")
        dir.create(path = embeddingsDir, recursive = TRUE)
        for (fl in embeddings) {
          file.copy(
            from = fl,
            to = embeddingsDir,
            overwrite = TRUE
          )
        }
      }


      if (nPivots > 0) {
        pivotTablesDir <- file.path(tmpDir, "xl", "pivotTables")
        dir.create(path = pivotTablesDir, recursive = TRUE)

        pivotTablesRelsDir <-
          file.path(tmpDir, "xl", "pivotTables", "_rels")
        dir.create(path = pivotTablesRelsDir, recursive = TRUE)

        pivotCacheDir <- file.path(tmpDir, "xl", "pivotCache")
        dir.create(path = pivotCacheDir, recursive = TRUE)

        pivotCacheRelsDir <-
          file.path(tmpDir, "xl", "pivotCache", "_rels")
        dir.create(path = pivotCacheRelsDir, recursive = TRUE)

        for (i in seq_along(pivotTables)) {
          file.copy(
            from = pivotTables[i],
            to = file.path(pivotTablesDir, sprintf("pivotTable%s.xml", i)),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }

        for (i in seq_along(pivotDefinitions)) {
          file.copy(
            from = pivotDefinitions[i],
            to = file.path(pivotCacheDir, sprintf("pivotCacheDefinition%s.xml", i)),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }

        for (i in seq_along(pivotRecords)) {
          file.copy(
            from = pivotRecords[i],
            to = file.path(pivotCacheDir, sprintf("pivotCacheRecords%s.xml", i)),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }

        for (i in seq_along(pivotDefinitionsRels)) {
          file.copy(
            from = pivotDefinitionsRels[i],
            to = file.path(
              pivotCacheRelsDir,
              sprintf("pivotCacheDefinition%s.xml.rels", i)
            ),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }

        for (i in seq_along(pivotTables.xml.rels)) {
          write_file(
            body = pivotTables.xml.rels[[i]],
            fl = file.path(pivotTablesRelsDir, sprintf("pivotTable%s.xml.rels", i))
          )
        }
      }

      ## slicers
      if (nSlicers > 0) {
        slicersDir <- file.path(tmpDir, "xl", "slicers")
        dir.create(path = slicersDir, recursive = TRUE)

        slicerCachesDir <- file.path(tmpDir, "xl", "slicerCaches")
        dir.create(path = slicerCachesDir, recursive = TRUE)

        for (i in seq_along(slicers)) {
          if (nchar(slicers[i]) > 0) {
            file.copy(from = slicers[i], to = file.path(slicersDir, sprintf("slicer%s.xml", i)))
          }
        }



        for (i in seq_along(slicerCaches)) {
          write_file(
            body = slicerCaches[[i]],
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


      ## write app.xml
      write_file(
        head = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
        body = "<Application>Microsoft Excel</Application>",
        tail = "</Properties>",
        fl = file.path(docPropsDir, "app.xml")
      )

      ## write core.xml
      write_file(
        head = "",
        body = pxml(core),
        tail = "",
        fl = file.path(docPropsDir, "core.xml")
      )

      ## write workbook.xml.rels
      write_file(
        head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        body = pxml(workbook.xml.rels),
        tail = "</Relationships>",
        fl = file.path(xlrelsDir, "workbook.xml.rels")
      )

      ## write tables
      if (length(unlist(tables, use.names = FALSE)) > 0) {
        for (i in seq_along(unlist(tables, use.names = FALSE))) {
          if (!grepl("openxlsx_deleted", attr(tables, "tableName")[i], fixed = TRUE)) {
            write_file(
              body = pxml(unlist(tables, use.names = FALSE)[[i]]),
              fl = file.path(xlTablesDir, sprintf("table%s.xml", i + 2))
            )
            if (tables.xml.rels[[i]] != "") {
              write_file(
                body = tables.xml.rels[[i]],
                fl = file.path(xlTablesRelsDir, sprintf("table%s.xml.rels", i + 2))
              )
            }
          }
        }
      }


      ## write query tables
      if (length(queryTables) > 0) {
        xlqueryTablesDir <- file.path(tmpDir, "xl", "queryTables")
        dir.create(path = xlqueryTablesDir, recursive = TRUE)

        for (i in seq_along(queryTables)) {
          write_file(
            body = queryTables[[i]],
            fl = file.path(xlqueryTablesDir, sprintf("queryTable%s.xml", i))
          )
        }
      }

      ## connections
      if (length(connections) > 0) {
        write_file(body = connections, fl = file.path(xlDir, "connections.xml"))
      }

      ## externalLinks
      if (length(externalLinks)) {
        externalLinksDir <- file.path(tmpDir, "xl", "externalLinks")
        dir.create(path = externalLinksDir, recursive = TRUE)

        for (i in seq_along(externalLinks)) {
          write_file(
            body = externalLinks[[i]],
            fl = file.path(externalLinksDir, sprintf("externalLink%s.xml", i))
          )
        }
      }

      ## externalLinks rels
      if (length(externalLinksRels)) {
        externalLinksRelsDir <-
          file.path(tmpDir, "xl", "externalLinks", "_rels")
        dir.create(path = externalLinksRelsDir, recursive = TRUE)

        for (i in seq_along(externalLinksRels)) {
          write_file(
            body = externalLinksRels[[i]],
            fl = file.path(
              externalLinksRelsDir,
              sprintf("externalLink%s.xml.rels", i)
            )
          )
        }
      }

      # # printerSettings
      # printDir <- file.path(tmpDir, "xl", "printerSettings")
      # dir.create(path = printDir, recursive = TRUE)
      # for (i in 1:nSheets) {
      #   writeLines(genPrinterSettings(), file.path(printDir, sprintf("printerSettings%s.bin", i)))
      # }

      ## media (copy file from origin to destination)
      for (x in media) {
        file.copy(x, file.path(xlmediaDir, names(media)[which(media == x)]))
      }

      ## VBA Macro
      if (!is.null(vbaProject)) {
        file.copy(vbaProject, xlDir)
      }

      ## write worksheet, worksheet_rels, drawings, drawing_rels
      .self$writeSheetDataXML(
        xldrawingsDir,
        xldrawingsRelsDir,
        xlworksheetsDir,
        xlworksheetsRelsDir
      )

      ## write sharedStrings.xml
      ct <- Content_Types
      if (length(sharedStrings) > 0) {
        write_file(
          head = sprintf(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%s" uniqueCount="%s">',
            length(sharedStrings),
            attr(sharedStrings, "uniqueCount")
          ),
          #body = stri_join(set_sst(attr(sharedStrings, "text")), collapse = "", sep = " "),
          body = stri_join(sharedStrings, collapse = "", sep = " "),
          tail = "</sst>",
          fl = file.path(xlDir, "sharedStrings.xml")
        )
      } else {
        ## Remove relationship to sharedStrings
        ct <- ct[!grepl("sharedStrings", ct)]
      }

      if (nComments > 0) {
        ct <-
          c(
            ct,
            '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
          )
      }

      ## write [Content_type]
      write_file(
        head = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        body = pxml(ct),
        tail = "</Types>",
        fl = file.path(tmpDir, "[Content_Types].xml")
      )


      styleXML <- styles
      styleXML$numFmts <-
        stri_join(
          sprintf('<numFmts count="%s">', length(styles$numFmts)),
          pxml(styles$numFmts),
          "</numFmts>"
        )
      styleXML$fonts <-
        stri_join(
          sprintf('<fonts count="%s">', length(styles$fonts)),
          pxml(styles$fonts),
          "</fonts>"
        )
      styleXML$fills <-
        stri_join(
          sprintf('<fills count="%s">', length(styles$fills)),
          pxml(styles$fills),
          "</fills>"
        )
      styleXML$borders <-
        stri_join(
          sprintf('<borders count="%s">', length(styles$borders)),
          pxml(styles$borders),
          "</borders>"
        )
      styleXML$cellStyleXfs <-
        c(
          sprintf('<cellStyleXfs count="%s">', length(styles$cellStyleXfs)),
          pxml(styles$cellStyleXfs),
          "</cellStyleXfs>"
        )
      styleXML$cellXfs <-
        stri_join(
          sprintf('<cellXfs count="%s">', length(styles$cellXfs)),
          paste0(styles$cellXfs, collapse = ""),
          "</cellXfs>"
        )
      styleXML$cellStyles <-
        stri_join(
          sprintf('<cellStyles count="%s">', length(styles$cellStyles)),
          pxml(styles$cellStyles),
          "</cellStyles>"
        )
      styleXML$dxfs <-
        ifelse(
          length(styles$dxfs) == 0,
          '<dxfs count="0"/>',
          stri_join(
            sprintf('<dxfs count="%s">', length(styles$dxfs)),
            stri_join(unlist(styles$dxfs), sep = " ", collapse = ""),
            "</dxfs>"
          )
        )

      ## write styles.xml
      if(class(styles_xml) == "uninitializedField") {
        write_file(
          head = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">',
          body = pxml(styleXML),
          tail = "</styleSheet>",
          fl = file.path(xlDir, "styles.xml")
        )
      } else {
        write_file(
          head = '',
          body = styles_xml,
          tail = '',
          fl = file.path(xlDir, "styles.xml")
        )
      }

      ## write workbook.xml
      workbookXML <- workbook
      workbookXML$sheets <-
        stri_join("<sheets>", pxml(workbookXML$sheets), "</sheets>")
      if (length(workbookXML$definedNames) > 0) {
        workbookXML$definedNames <-
          stri_join(
            "<definedNames>",
            pxml(workbookXML$definedNames),
            "</definedNames>"
          )
      }

      write_file(
        head = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">',
        body = pxml(workbookXML),
        tail = "</workbook>",
        fl = file.path(xlDir, "workbook.xml")
      )
      .self$workbook$sheets <-
        workbook$sheets[order(sheetOrder)] ## Need to reset sheet order to allow multiple savings

      ## compress to xlsx
      wd <- getwd()
      tmpFile <-
        basename(tempfile(fileext = ifelse(is.null(vbaProject), ".xlsx", ".xlsm")))
      on.exit(expr = setwd(wd), add = TRUE)

      ## zip it
      setwd(dir = tmpDir)
      cl <-
        ifelse(
          !is.null(getOption("openxlsx.compresssionLevel")),
          getOption("openxlsx.compresssionLevel"),
          getOption("openxlsx.compresssionevel", 6)
        )
      zipr(
        zipfile = tmpFile, include_directories = FALSE,
        files = list.files(path = tmpDir, all.files = FALSE),
        recurse = TRUE,
        compression_level = cl
      )

      # reset styles - maintain any changes to base font
      # TODO: why would I want to do that?
      baseFont <- styles$fonts[[1]]
      .self$styles <-
        genBaseStyleSheet(styles$dxfs,
          tableStyles = styles$tableStyles,
          extLst = styles$extLst
        )
      .self$styles$fonts[[1]] <- baseFont


      return(file.path(tmpDir, tmpFile))
    },

    updateSharedStrings = function(uNewStr) {
      ## Function will return named list of references to new strings
      uStr <- uNewStr[which(!uNewStr %in% sharedStrings)]
      uCount <- attr(sharedStrings, "uniqueCount")
      .self$sharedStrings <- append(sharedStrings, uStr)

      attr(.self$sharedStrings, "uniqueCount") <- uCount + length(uStr)
    },

    validateSheet = function(sheetName) {
      if (!is.numeric(sheetName)) {
        if (is.null(sheet_names)) {
          stop("Workbook does not contain any worksheets.", call. = FALSE)
        }
      }

      if (is.numeric(sheetName)) {
        if (sheetName > length(sheet_names)) {
          stop(sprintf("This Workbook only has %s sheets.", length(sheet_names)),
            call. =
              FALSE
          )
        }

        return(sheetName)
      } else if (!sheetName %in% replaceXMLEntities(sheet_names)) {
        stop(sprintf("Sheet '%s' does not exist.", replaceXMLEntities(sheetName)), call. = FALSE)
      }

      return(which(replaceXMLEntities(sheet_names) == sheetName))
    },

    getSheetName = function(sheetIndex) {
      if (any(length(sheet_names) < sheetIndex)) {
        stop(sprintf("Workbook only contains %s sheet(s).", length(sheet_names)))
      }

      sheet_names[sheetIndex]
    },

    buildTable = function(sheet,
      colNames,
      ref,
      showColNames,
      tableStyle,
      tableName,
      withFilter,
      totalsRowCount = 0,
      showFirstColumn = 0,
      showLastColumn = 0,
      showRowStripes = 1,
      showColumnStripes = 0) {
      ## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
      id <- as.character(length(tables) + 3L)
      sheet <- validateSheet(sheet)

      ## build table XML and save to tables field
      table <-
        sprintf(
          '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="%s" name="%s" displayName="%s" ref="%s" totalsRowCount="%s"',
          id,
          tableName,
          tableName,
          ref,
          as.integer(totalsRowCount)
        )
      # because tableName might be native encoded non-ASCII strings, we need to ensure
      # it's UTF-8 encoded
      table <- enc2utf8(table)

      nms <- names(tables)
      tSheets <- attr(tables, "sheet")
      tNames <- attr(tables, "tableName")

      tableStyleXML <-
        sprintf(
          '<tableStyleInfo name="%s" showFirstColumn="%s" showLastColumn="%s" showRowStripes="%s" showColumnStripes="%s"/>',
          tableStyle,
          as.integer(showFirstColumn),
          as.integer(showLastColumn),
          as.integer(showRowStripes),
          as.integer(showColumnStripes)
        )


      .self$tables <-
        c(
          tables,
          build_table_xml(
            table = table,
            tableStyleXML = tableStyleXML,
            ref = ref,
            colNames = gsub("\n|\r", "_x000a_", colNames),
            showColNames = showColNames,
            withFilter = withFilter
          )
        )
      names(.self$tables) <- c(nms, ref)
      attr(.self$tables, "sheet") <- c(tSheets, sheet)
      attr(.self$tables, "tableName") <- c(tNames, tableName)

      .self$worksheets[[sheet]]$tableParts <-
        append(
          worksheets[[sheet]]$tableParts,
          sprintf('<tablePart r:id="rId%s"/>', id)
        )
      attr(.self$worksheets[[sheet]]$tableParts, "tableName") <-
        c(tNames[tSheets == sheet &
            !grepl("openxlsx_deleted", tNames, fixed = TRUE)], tableName)



      ## update Content_Types
      .self$Content_Types <-
        c(
          Content_Types,
          sprintf(
            '<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
            id
          )
        )

      ## create a table.xml.rels
      .self$tables.xml.rels <- append(tables.xml.rels, "")

      ## update worksheets_rels
      .self$worksheets_rels[[sheet]] <- c(
        worksheets_rels[[sheet]],
        sprintf(
          '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',
          id,
          id
        )
      )
    },

    writeDrawingVML = function(dir) {
      for (i in seq_along(comments)) {
        id <- 1025

        cd <- unlist(lapply(comments[[i]], "[[", "clientData"))
        nComments <- length(cd)

        ## write head
        if (nComments > 0 | length(vml[[i]]) > 0) {
          write(
            x = stri_join(
              '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
            ),
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            sep = " "
          )
        }

        if (nComments > 0) {
          for (j in 1:nComments) {
            id <- id + 1L
            write(
              x = genBaseShapeVML(cd[j], id),
              file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
              append = TRUE
            )
          }
        }

        if (length(vml[[i]]) > 0) {
          write(
            x = vml[[i]],
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            append = TRUE
          )
        }

        if (nComments > 0 | length(vml[[i]]) > 0) {
          write(
            x = "</xml>",
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            append = TRUE
          )
          .self$worksheets[[i]]$legacyDrawing <-
            '<legacyDrawing r:id="rId2"/>'
        }

      }

      for (i in seq_along(drawings_vml)) {
        write(
          x = drawings_vml[[i]],
          file = file.path(dir, sprintf("vmlDrawing%s.vml", i))
        )
      }
    },

    updateStyles = function(style) {
      ## Updates styles.xml
      xfNode <- list(
        numFmtId = 0,
        fontId = 0,
        fillId = 0,
        borderId = 0,
        xfId = 0
      )


      alignmentFlag <- FALSE

      ## Font
      if (!is.null(style$fontName) |
          !is.null(style$fontSize) |
          !is.null(style$fontColour) |
          !is.null(style$fontDecoration) |
          !is.null(style$fontFamily) |
          !is.null(style$fontScheme)) {
        fontNode <- .self$createFontNode(style)
        fontId <- style$fontId

        if (length(fontId) == 0) {
          fontId <- style$fontId
          .self$styles$fonts <- append(styles[["fonts"]], fontNode)
        }

        xfNode$fontId <- fontId
        xfNode <- append(xfNode, list("applyFont" = "1"))
      }


      ## numFmt
      if (!is.null(style$numFmt)) {
        if (as.integer(style$numFmt$numFmtId) > 0) {
          numFmtId <- style$numFmt$numFmtId
          if (as.integer(numFmtId) > 163L) {
            tmp <- style$numFmt$formatCode

            .self$styles$numFmts <- unique(c(
              styles$numFmts,
              sprintf(
                '<numFmt numFmtId="%s" formatCode="%s"/>',
                numFmtId,
                tmp
              )
            ))
          }

          xfNode$numFmtId <- numFmtId
          xfNode <- append(xfNode, list("applyNumberFormat" = "1"))
        }
      }

      ## Fill
      if (!is.null(style$fill)) {
        fillNode <- .self$createFillNode(style)
        if (!is.null(fillNode)) {
          fillId <- which(styles$fills == fillNode) - 1L

          if (length(fillId) == 0) {
            fillId <- length(styles$fills)
            .self$styles$fills <- c(styles$fills, fillNode)
          }
          xfNode$fillId <- fillId
          xfNode <- append(xfNode, list("applyFill" = "1"))
        }
      }

      ## Border
      if (any(!is.null(
        c(
          style$borderLeft,
          style$borderRight,
          style$borderTop,
          style$borderBottom,
          style$borderDiagonal
        )
      ))) {
        borderNode <- .self$createBorderNode(style)
        borderId <- which(styles$borders == borderNode) - 1L

        if (length(borderId) == 0) {
          borderId <- length(styles$borders)
          .self$styles$borders <- c(styles$borders, borderNode)
        }

        xfNode$borderId <- borderId
        xfNode <- append(xfNode, list("applyBorder" = "1"))
      }


      if(!is.null(style$xfId))
        xfNode$xfId <- style$xfId

      childNodes <- ""

      ## Alignment
      if (!is.null(style$halign) |
          !is.null(style$valign) |
          !is.null(style$wrapText) |
          !is.null(style$textRotation) | !is.null(style$indent)) {
        attrs <- list()
        alignNode <- "<alignment"

        if (!is.null(style$textRotation)) {
          alignNode <-
            stri_join(alignNode,
              sprintf('textRotation="%s"', style$textRotation),
              sep = " "
            )
        }

        if (!is.null(style$halign)) {
          alignNode <-
            stri_join(alignNode, sprintf('horizontal="%s"', style$halign), sep = " ")
        }

        if (!is.null(style$valign)) {
          alignNode <-
            stri_join(alignNode, sprintf('vertical="%s"', style$valign), sep = " ")
        }

        if (!is.null(style$indent)) {
          alignNode <-
            stri_join(alignNode, sprintf('indent="%s"', style$indent), sep = " ")
        }

        if (!is.null(style$wrapText)) {
          if (style$wrapText) {
            alignNode <- stri_join(alignNode, 'wrapText="1"', sep = " ")
          }
        }


        alignNode <- stri_join(alignNode, "/>")

        alignmentFlag <- TRUE
        xfNode <- append(xfNode, list("applyAlignment" = "1"))

        childNodes <- stri_join(childNodes, alignNode)
      }

      if (!is.null(style$hidden) | !is.null(style$locked)) {
        xfNode <- append(xfNode, list("applyProtection" = "1"))
        protectionNode <- "<protection"

        if (!is.null(style$hidden)) {
          protectionNode <-
            stri_join(protectionNode, sprintf('hidden="%s"', as.numeric(style$hidden)), sep = " ")
        }
        if (!is.null(style$locked)) {
          protectionNode <-
            stri_join(protectionNode, sprintf('locked="%s"', as.numeric(style$locked)), sep = " ")
        }

        protectionNode <- stri_join(protectionNode, "/>")
        childNodes <- stri_join(childNodes, protectionNode)
      }

      if (length(childNodes) > 0) {
        xfNode <-
          stri_join(
            "<xf ",
            stri_join(
              stri_join(names(xfNode), '="', xfNode, '"'),
              sep = " ",
              collapse = " "
            ),
            ">",
            childNodes,
            "</xf>"
          )
      } else {
        xfNode <-
          stri_join("<xf ", stri_join(
            stri_join(names(xfNode), '="', xfNode, '"'),
            sep = " ",
            collapse = " "
          ), "/>")
      }

      styleId <- which(styles$cellXfs == xfNode) - 1L
      if (length(styleId) == 0) {
        styleId <- length(styles$cellXfs)
        .self$styles$cellXfs <- c(styles$cellXfs, xfNode)
      }


      return(as.integer(styleId))
    },

    updateCellStyles = function() {
      flag <- TRUE
      for (style in cellStyleObjects) {
        ## Updates styles.xml
        xfNode <- list(
          numFmtId = 0,
          fontId = 0,
          fillId = 0,
          borderId = 0
        )


        alignmentFlag <- FALSE

        ## Font
        if (!is.null(style$fontName) |
            !is.null(style$fontSize) |
            !is.null(style$fontColour) |
            !is.null(style$fontDecoration) |
            !is.null(style$fontFamily) |
            !is.null(style$fontScheme)) {
          fontNode <- .self$createFontNode(style)
          fontId <- which(styles$font == fontNode) - 1L

          if (length(fontId) == 0) {
            fontId <- length(styles$fonts)
            .self$styles$fonts <- append(styles[["fonts"]], fontNode)
          }

          xfNode$fontId <- fontId
          xfNode <- append(xfNode, list("applyFont" = "1"))
        }


        ## numFmt
        if (!is.null(style$numFmt)) {
          if (as.integer(style$numFmt$numFmtId) > 0) {
            numFmtId <- style$numFmt$numFmtId
            if (as.integer(numFmtId) > 163L) {
              tmp <- style$numFmt$formatCode

              .self$styles$numFmts <- unique(c(
                styles$numFmts,
                sprintf(
                  '<numFmt numFmtId="%s" formatCode="%s"/>',
                  numFmtId,
                  tmp
                )
              ))
            }

            xfNode$numFmtId <- numFmtId
            xfNode <- append(xfNode, list("applyNumberFormat" = "1"))
          }
        }

        ## Fill
        if (!is.null(style$fill)) {
          fillNode <- .self$createFillNode(style)
          if (!is.null(fillNode)) {
            fillId <- which(styles$fills == fillNode) - 1L

            if (length(fillId) == 0) {
              fillId <- length(styles$fills)
              .self$styles$fills <- c(styles$fills, fillNode)
            }
            xfNode$fillId <- fillId
            xfNode <- append(xfNode, list("applyFill" = "1"))
          }
        }

        ## Border
        if (any(!is.null(
          c(
            style$borderLeft,
            style$borderRight,
            style$borderTop,
            style$borderBottom,
            style$borderDiagonal
          )
        ))) {
          borderNode <- .self$createBorderNode(style)
          borderId <- which(styles$borders == borderNode) - 1L

          if (length(borderId) == 0) {
            borderId <- length(styles$borders)
            .self$styles$borders <- c(styles$borders, borderNode)
          }

          xfNode$borderId <- borderId
          xfNode <- append(xfNode, list("applyBorder" = "1"))
        }

        xfNode <-
          stri_join("<xf ", stri_join(
            stri_join(names(xfNode), '="', xfNode, '"'),
            sep = " ",
            collapse = " "
          ), "/>")

        if (flag) {
          .self$styles$cellStyleXfs <- xfNode
          flag <- FALSE
        } else {
          .self$styles$cellStyleXfs <- c(styles$cellStyleXfs, xfNode)
        }
      }
    },

    createFontNode = function(style) {
      baseFont <- .self$getBaseFont()

      fontNode <- "<font>"

      ## size
      if (is.null(style$fontSize[[1]])) {
        fontNode <-
          stri_join(fontNode, sprintf('<sz %s="%s"/>', names(baseFont$size), baseFont$size))
      } else {
        fontNode <-
          stri_join(fontNode, sprintf('<sz %s="%s"/>', names(style$fontSize), style$fontSize))
      }

      ## colour
      if (is.null(style$fontColour[[1]])) {
        fontNode <-
          stri_join(
            fontNode,
            sprintf(
              '<color %s="%s"/>',
              names(baseFont$colour),
              baseFont$colour
            )
          )
      } else {
        if (length(style$fontColour) > 1) {
          fontNode <- stri_join(fontNode, sprintf(
            "<color %s/>",
            stri_join(
              sapply(seq_along(style$fontColour), function(i) {
                sprintf('%s="%s"', names(style$fontColour)[i], style$fontColour[i])
              }),
              sep = " ",
              collapse = " "
            )
          ))
        } else {
          fontNode <-
            stri_join(
              fontNode,
              sprintf(
                '<color %s="%s"/>',
                names(style$fontColour),
                style$fontColour
              )
            )
        }
      }


      ## name
      if (is.null(style$fontName[[1]])) {
        fontNode <-
          stri_join(
            fontNode,
            sprintf('<name %s="%s"/>', names(baseFont$name), baseFont$name)
          )
      } else {
        fontNode <-
          stri_join(
            fontNode,
            sprintf('<name %s="%s"/>', names(style$fontName), style$fontName)
          )
      }

      ### Create new font and return Id
      if (!is.null(style$fontFamily)) {
        fontNode <-
          stri_join(fontNode, sprintf('<family val = "%s"/>', style$fontFamily))
      }

      if (!is.null(style$fontScheme)) {
        fontNode <-
          stri_join(fontNode, sprintf('<scheme val = "%s"/>', style$fontScheme))
      }

      if ("BOLD" %in% style$fontDecoration) {
        fontNode <- stri_join(fontNode, "<b/>")
      }

      if ("ITALIC" %in% style$fontDecoration) {
        fontNode <- stri_join(fontNode, "<i/>")
      }

      if ("UNDERLINE" %in% style$fontDecoration) {
        fontNode <- stri_join(fontNode, '<u val="single"/>')
      }

      if ("UNDERLINE2" %in% style$fontDecoration) {
        fontNode <- stri_join(fontNode, '<u val="double"/>')
      }

      if ("STRIKEOUT" %in% style$fontDecoration) {
        fontNode <- stri_join(fontNode, "<strike/>")
      }

      stri_join(fontNode, "</font>")
    },

    getBaseFont = function() {
      baseFont <- styles$fonts[[1]]

      sz     <- font_val(baseFont, "font", "sz")
      colour <- font_val(baseFont, "font", "color")
      name   <- font_val(baseFont, "font", "name")

      if (length(sz[[1]]) == 0) {
        sz <- list("val" = "10")
      }

      if (length(colour[[1]]) == 0) {
        colour <- list("rgb" = "#000000")
      }

      if (length(name[[1]]) == 0) {
        name <- list("val" = "Calibri")
      }

      list(
        "size" = sz,
        "colour" = colour,
        "name" = name
      )
    },

    createBorderNode = function(style) {
      borderNode <- "<border"

      if (style$borderDiagonalUp) {
        borderNode <- stri_join(borderNode, 'diagonalUp="1"', sep = " ")
      }

      if (style$borderDiagonalDown) {
        borderNode <-
          stri_join(borderNode, 'diagonalDown="1"', sep = " ")
      }

      borderNode <- stri_join(borderNode, ">")

      if (!is.null(style$borderLeft)) {
        borderNode <-
          stri_join(
            borderNode,
            sprintf('<left style="%s">', style$borderLeft),
            sprintf(
              '<color %s="%s"/>',
              names(style$borderLeftColour),
              style$borderLeftColour
            ),
            "</left>"
          )
      }

      if (!is.null(style$borderRight)) {
        borderNode <-
          stri_join(
            borderNode,
            sprintf('<right style="%s">', style$borderRight),
            sprintf(
              '<color %s="%s"/>',
              names(style$borderRightColour),
              style$borderRightColour
            ),
            "</right>"
          )
      }

      if (!is.null(style$borderTop)) {
        borderNode <-
          stri_join(
            borderNode,
            sprintf('<top style="%s">', style$borderTop),
            sprintf(
              '<color %s="%s"/>',
              names(style$borderTopColour),
              style$borderTopColour
            ),
            "</top>"
          )
      }

      if (!is.null(style$borderBottom)) {
        borderNode <-
          stri_join(
            borderNode,
            sprintf('<bottom style="%s">', style$borderBottom),
            sprintf(
              '<color %s="%s"/>',
              names(style$borderBottomColour),
              style$borderBottomColour
            ),
            "</bottom>"
          )
      }

      if (!is.null(style$borderDiagonal)) {
        borderNode <-
          stri_join(
            borderNode,
            sprintf('<diagonal style="%s">', style$borderDiagonal),
            sprintf(
              '<color %s="%s"/>',
              names(style$borderDiagonalColour),
              style$borderDiagonalColour
            ),
            "</diagonal>"
          )
      }

      stri_join(borderNode, "</border>")
    },

    createFillNode = function(style, patternType = "solid") {
      fill <- style$fill

      ## gradientFill
      if (any(grepl("gradientFill", fill))) {
        fillNode <- fill # stri_join("<fill>", fill, "</fill>")
      } else if (!is.null(fill$fillFg) | !is.null(fill$fillBg)) {
        fillNode <-
          stri_join(
            "<fill>",
            sprintf('<patternFill patternType="%s">', patternType)
          )

        if (!is.null(fill$fillFg)) {
          fillNode <-
            stri_join(fillNode, sprintf(
              "<fgColor %s/>",
              stri_join(
                stri_join(names(fill$fillFg), '="', fill$fillFg, '"'),
                sep = " ",
                collapse = " "
              )
            ))
        }

        if (!is.null(fill$fillBg)) {
          fillNode <-
            stri_join(fillNode, sprintf(
              "<bgColor %s/>",
              stri_join(
                stri_join(names(fill$fillBg), '="', fill$fillBg, '"'),
                sep = " ",
                collapse = " "
              )
            ))
        }

        fillNode <- stri_join(fillNode, "</patternFill></fill>")
      } else {
        return(NULL)
      }

      return(fillNode)
    },

    setSheetName = function(sheet, newSheetName) {
      if (newSheetName %in% sheet_names) {
        stop(sprintf("Sheet %s already exists!", newSheetName))
      }

      sheet <- validateSheet(sheet)

      oldName <- sheet_names[[sheet]]
      .self$sheet_names[[sheet]] <- newSheetName

      ## Rename in workbook
      sheetId <-
        regmatches(
          workbook$sheets[[sheet]],
          regexpr('(?<=sheetId=")[0-9]+', workbook$sheets[[sheet]], perl = TRUE)
        )
      rId <-
        regmatches(
          workbook$sheets[[sheet]],
          regexpr('(?<= r:id="rId)[0-9]+', workbook$sheets[[sheet]], perl = TRUE)
        )
      .self$workbook$sheets[[sheet]] <-
        sprintf(
          '<sheet name="%s" sheetId="%s" r:id="rId%s"/>',
          newSheetName,
          sheetId,
          rId
        )

      ## rename styleObjects sheet component
      if (length(styleObjects) > 0) {
        .self$styleObjects <- lapply(styleObjects, function(x) {
          if (x$sheet == oldName) {
            x$sheet <- newSheetName
          }

          return(x)
        })
      }

      ## rename defined names
      if (length(workbook$definedNames) > 0) {
        belongTo <- getDefinedNamesSheet(workbook$definedNames)
        toChange <- belongTo == oldName
        if (any(toChange)) {
          newSheetName <- sprintf("'%s'", newSheetName)
          tmp <-
            gsub(oldName, newSheetName, workbook$definedName[toChange], fixed = TRUE)
          tmp <- gsub("'+", "'", tmp)
          .self$workbook$definedNames[toChange] <- tmp
        }
      }
    },

    writeSheetDataXML = function(xldrawingsDir,
      xldrawingsRelsDir,
      xlworksheetsDir,
      xlworksheetsRelsDir) {
      ## write worksheets
      nSheets <- length(worksheets)

      for (i in seq_len(nSheets)) {
        ## Write drawing i (will always exist) skip those that are empty
        if (!identical(drawings[[i]], list())) {
          write_file(
            head = '',
            body = pxml(drawings[[i]]),
            tail = '',
            fl = file.path(xldrawingsDir, stri_join("drawing", i, ".xml"))
          )
          if (!identical(drawings_rels[[i]], list())) {
            write_file(
              head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
              body = pxml(drawings_rels[[i]]),
              tail = "</Relationships>",
              fl = file.path(xldrawingsRelsDir, stri_join("drawing", i, ".xml.rels"))
            )
          }
        } else {
          .self$worksheets[[i]]$drawing <- character(0)
        }

        ## vml drawing
        if (length(vml_rels[[i]]) > 0) {
          file.copy(
            from = vml_rels[[i]],
            to = file.path(
              xldrawingsRelsDir,
              stri_join("vmlDrawing", i, ".vml.rels")
            )
          )
        }

        # outlineLevelRow in SheetformatPr
        if ((length(outlineLevels[[i]]) > 0) && (!grepl("outlineLevelRow", worksheets[[i]]$sheetFormatPr))) {
          .self$worksheets[[i]]$sheetFormatPr <- gsub("/>", ' outlineLevelRow="1"/>', worksheets[[i]]$sheetFormatPr)
        }

        if (isChartSheet[i]) {
          chartSheetDir <- file.path(dirname(xlworksheetsDir), "chartsheets")
          chartSheetRelsDir <-
            file.path(dirname(xlworksheetsDir), "chartsheets", "_rels")

          if (!file.exists(chartSheetDir)) {
            dir.create(chartSheetDir, recursive = TRUE)
            dir.create(chartSheetRelsDir, recursive = TRUE)
          }

          write_file(
            body = worksheets[[i]]$get_prior_sheet_data(),
            fl = file.path(chartSheetDir, stri_join("sheet", i, ".xml"))
          )

          write_file(
            head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            body = pxml(worksheets_rels[[i]]),
            tail = "</Relationships>",
            fl = file.path(chartSheetRelsDir, sprintf("sheet%s.xml.rels", i))
          )
        } else {
          ## Write worksheets
          ws <- worksheets[[i]]
          hasHL <-
            ifelse(length(ws$hyperlinks) > 0, TRUE, FALSE)

          ## reorder sheet data
          ws$order_sheetdata()

          prior <- ws$get_prior_sheet_data()
          post <- ws$get_post_sheet_data()

          # worksheets[[i]]$sheet_data$style_id <-
          #   as.character(worksheets[[i]]$sheet_data$style_id)

          cc <- ws$sheet_data$cc


          cc$r <- paste0(cc$c_r, cc$row_r)
          # prepare data for output

          # there can be files, where row_attr is incomplete because a row
          # is lacking any attributes (presumably was added before saving)
          # still row_attr is what we want!
          cc_rows <- names(ws$sheet_data$row_attr)
          cc_out <- vector("list", length = length(cc_rows))
          names(cc_out) <- cc_rows

          for (cc_r in cc_rows) {
            tmp <- cc[cc$row_r == cc_r, c("r", "v", "c_t", "c_s", "f", "f_t", "f_ref", "f_si", "is")]
            nams <- cc[cc$row_r == cc_r, c("c_r")]
            ltmp <- vector("list", nrow(tmp))
            names(ltmp) <- nams

            for (nr in seq_len(nrow(tmp))) {
              ltmp[[nr]] <- as.list(tmp[nr, ])
            }

            cc_out[[cc_r]] <- ltmp
          }

          ws$sheet_data$cc_out <- cc_out

          # row_attr <- ws$sheet_data$row_attr
          # nam_at <- names(row_attr)
          # wanted <- as.character(seq(min(as.numeric(nam_at)),
          #                            max(as.numeric(nam_at))))
          # empty_row_attr <- wanted[!wanted %in% nam_at]
          # # add empty list
          # if(!identical(empty_row_attr, character(0)))
          #   row_attr[[empty_row_attr]] <- list()
          # # restore order
          # ws$sheet_data$row_attr <- row_attr[wanted]

          # message(i, " \n")
          write_worksheet_xml_2(
            prior = prior,
            post = post,
            sheet_data = ws$sheet_data,
            cols_attr = ws$cols_attr,
            rows_attr = ws$sheet_data$row_attr,
            row_heights_ = NULL,
            outline_levels_ = unlist(outlineLevels[[i]]),
            R_fileName = file.path(xlworksheetsDir, sprintf("sheet%s.xml", i))
          )

          # # why would I want to erase everything in here?
          # worksheets[[i]]$sheet_data$style_id <- integer(0)


          ## write worksheet rels
          if (length(worksheets_rels[[i]]) > 0) {
            ws_rels <- worksheets_rels[[i]]
            if (hasHL) {
              h_inds <- stri_join(seq_along(worksheets[[i]]$hyperlinks), "h")
              ws_rels <-
                c(ws_rels, unlist(
                  lapply(seq_along(h_inds), function(j) {
                    worksheets[[i]]$hyperlinks[[j]]$to_target_xml(h_inds[j])
                  })
                ))
            }

            ## Check if any tables were deleted - remove these from rels
            if (length(tables) > 0) {
              table_inds <- which(grepl("tables/table[0-9].xml", ws_rels))

              if (length(table_inds) > 0) {
                ids <-
                  regmatches(
                    ws_rels[table_inds],
                    regexpr(
                      '(?<=Relationship Id=")[0-9A-Za-z]+',
                      ws_rels[table_inds],
                      perl = TRUE
                    )
                  )
                inds <-
                  as.integer(gsub("[^0-9]", "", ids, perl = TRUE)) - 2L
                table_nms <- attr(tables, "tableName")[inds]
                is_deleted <-
                  grepl("openxlsx_deleted", table_nms, fixed = TRUE)
                if (any(is_deleted)) {
                  ws_rels <- ws_rels[-table_inds[is_deleted]]
                }
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
      } ## end of loop through 1:nSheets

      invisible(0)
    },


    setRowHeights = function(sheet, rows, heights) {
      sheet <- validateSheet(sheet)

      ## remove any conflicting heights
      flag <- names(rowHeights[[sheet]]) %in% rows
      if (any(flag)) {
        .self$rowHeights[[sheet]] <- rowHeights[[sheet]][!flag]
      }

      nms <- c(names(rowHeights[[sheet]]), rows)
      allRowHeights <- unlist(c(rowHeights[[sheet]], heights))
      names(allRowHeights) <- nms

      allRowHeights <-
        allRowHeights[order(as.integer(names(allRowHeights)))]

      .self$rowHeights[[sheet]] <- allRowHeights
    },

    groupColumns = function(sheet) {
      sheet <- validateSheet(sheet)

      hidden <- attr(colOutlineLevels[[sheet]], "hidden", exact = TRUE)
      cols <- names(colOutlineLevels[[sheet]])

      if (!grepl("outlineLevelCol", worksheets[[sheet]]$sheetFormatPr)) {
        .self$worksheets[[sheet]]$sheetFormatPr <- sub("/>", ' outlineLevelCol="1"/>', worksheets[[sheet]]$sheetFormatPr)
      }

      # Check if column is already created (by `setColWidths()` or on import)
      # Note that columns are initiated by `setColWidths` first (see: order of execution in `preSaveCleanUp()`)
      if (any(cols %in% names(worksheets[[sheet]]$cols))) {

        for (i in intersect(cols, names(worksheets[[sheet]]$cols))) {
          outline_hidden <- attr(colOutlineLevels[[sheet]], "hidden")[attr(colOutlineLevels[[sheet]], "names") == i]

          if (grepl("outlineLevel", worksheets[[sheet]]$cols[[i]], perl = TRUE)) {
            .self$worksheets[[sheet]]$cols[[i]] <- sub("((?<=hidden=\")(\\w+)\")", paste0(outline_hidden, "\""), worksheets[[sheet]]$cols[[i]], perl = TRUE)
          } else {
            .self$worksheets[[sheet]]$cols[[i]] <- sub("((?<=hidden=\")(\\w+)\")", paste0(outline_hidden, "\" outlineLevel=\"1\""), worksheets[[sheet]]$cols[[i]], perl = TRUE)
          }
        }

        cols <- cols[!cols %in% names(worksheets[[sheet]]$cols)]
        hidden <- attr(colOutlineLevels[[sheet]], "hidden")[attr(colOutlineLevels[[sheet]], "names") %in% cols]
      }

      if (length(cols) > 0) {
        colNodes <- sprintf('<col min="%s" max="%s" outlineLevel="1" hidden="%s"/>', cols, cols, hidden)
        names(colNodes) <- cols
        .self$worksheets[[sheet]]$cols <- append(worksheets[[sheet]]$cols, colNodes)
      }
    },

    groupRows = function(sheet, rows, hidden, levels) {
      sheet <- validateSheet(sheet)


      flag <- names(outlineLevels[[sheet]]) %in% rows
      if (any(flag)) {
        .self$outlineLevels[[sheet]] <- outlineLevels[[sheet]][!flag]
      }

      nms <- c(names(outlineLevels[[sheet]]), rows)

      allOutlineLevels <- unlist(c(outlineLevels[[sheet]], levels))
      names(allOutlineLevels) <- nms

      existing_hidden <- attr(outlineLevels[[sheet]], "hidden", exact = TRUE)
      all_hidden <- c(existing_hidden, as.character(as.integer(hidden)))

      allOutlineLevels <-
        allOutlineLevels[order(as.integer(names(allOutlineLevels)))]

      .self$outlineLevels[[sheet]] <- allOutlineLevels

      attr(.self$outlineLevels[[sheet]], "hidden") <- as.character(as.integer(all_hidden))


      if (!grepl("outlineLevelRow", worksheets[[sheet]]$sheetFormatPr)) {
        .self$worksheets[[sheet]]$sheetFormatPr <- gsub("/>", ' outlineLevelRow="1"/>', worksheets[[sheet]]$sheetFormatPr)
      }
    },

    deleteWorksheet = function(sheet) {
      # To delete a worksheet
      # Remove colwidths element
      # Remove drawing partname from Content_Types (drawing(sheet).xml)
      # Remove highest sheet from Content_Types
      # Remove drawings element
      # Remove drawings_rels element

      # Remove vml element
      # Remove vml_rels element

      # Remove rowHeights element
      # Remove styleObjects on sheet
      # Remove last sheet element from workbook
      # Remove last sheet element from workbook.xml.rels
      # Remove element from worksheets
      # Remove element from worksheets_rels
      # Remove hyperlinks
      # Reduce calcChain i attributes & remove calcs on sheet
      # Remove sheet from sheetOrder
      # Remove queryTable references from workbook$definedNames to worksheet
      # remove tables

      sheet <- validateSheet(sheet)
      sheetNames <- sheet_names
      nSheets <- length(unlist(sheetNames, use.names = FALSE))
      sheetName <- sheetNames[[sheet]]

      .self$colWidths[[sheet]] <- NULL
      .self$sheet_names <- sheet_names[-sheet]

      ## remove last drawings(sheet).xml from Content_Types
      .self$Content_Types <-
        Content_Types[!grepl(sprintf("drawing%s.xml", nSheets), Content_Types)]

      ## remove highest sheet
      .self$Content_Types <-
        Content_Types[!grepl(sprintf("sheet%s.xml", nSheets), Content_Types)]

      .self$drawings[[sheet]] <- NULL
      .self$drawings_rels[[sheet]] <- NULL

      .self$vml[[sheet]] <- NULL
      .self$vml_rels[[sheet]] <- NULL

      .self$rowHeights[[sheet]] <- NULL
      .self$colOutlineLevels[[sheet]] <- NULL
      .self$outlineLevels[[sheet]] <- NULL
      .self$comments[[sheet]] <- NULL
      .self$threadComments[[sheet]] <- NULL
      .self$isChartSheet <- isChartSheet[-sheet]

      ## sheetOrder
      toRemove <- which(sheetOrder == sheet)
      .self$sheetOrder[sheetOrder > sheet] <-
        sheetOrder[sheetOrder > sheet] - 1L
      .self$sheetOrder <- sheetOrder[-toRemove]


      ## remove styleObjects
      if (length(styleObjects) > 0) {
        .self$styleObjects <-
          styleObjects[unlist(lapply(styleObjects, "[[", "sheet"), use.names = FALSE) != sheetName]
      }

      ## Need to remove reference from workbook.xml.rels to pivotCache
      removeRels <-
        worksheets_rels[[sheet]][grepl("pivotTables", worksheets_rels[[sheet]])]
      if (length(removeRels) > 0) {
        ## sheet rels links to a pivotTable file, the corresponding pivotTable_rels file links to the cacheDefn which is listing in workbook.xml.rels
        ## remove reference to this file from the workbook.xml.rels
        fileNo <-
          as.integer(unlist(regmatches(
            removeRels,
            gregexpr("(?<=pivotTable)[0-9]+(?=\\.xml)", removeRels, perl = TRUE)
          )))
        toRemove <-
          stri_join(
            sprintf("(pivotCacheDefinition%s\\.xml)", fileNo),
            sep = " ",
            collapse = "|"
          )

        fileNo <- which(grepl(toRemove, pivotTables.xml.rels))
        toRemove <-
          stri_join(
            sprintf("(pivotCacheDefinition%s\\.xml)", fileNo),
            sep = " ",
            collapse = "|"
          )

        ## remove reference to file from workbook.xml.res
        .self$workbook.xml.rels <-
          workbook.xml.rels[!grepl(toRemove, workbook.xml.rels)]
      }

      ## As above for slicers
      ## Need to remove reference from workbook.xml.rels to pivotCache
      removeRels <- grepl("slicers", worksheets_rels[[sheet]])
      if (any(removeRels)) {
        .self$workbook.xml.rels <-
          workbook.xml.rels[!grepl(sprintf("(slicerCache%s\\.xml)", sheet), workbook.xml.rels)]
      }

      ## wont't remove tables and then won't need to reassign table r:id's but will rename them!
      .self$worksheets[[sheet]] <- NULL
      .self$worksheets_rels[[sheet]] <- NULL

      if (length(tables) > 0) {
        tableSheets <- attr(tables, "sheet")
        tableNames <- attr(tables, "tableName")

        inds <-
          tableSheets %in% sheet &
          !grepl("openxlsx_deleted", attr(tables, "tableName"), fixed = TRUE)
        tableSheets[tableSheets > sheet] <-
          tableSheets[tableSheets > sheet] - 1L

        ## Need to flag a table as deleted
        if (any(inds)) {
          tableSheets[inds] <- 0
          tableNames[inds] <-
            stri_join(tableNames[inds], "_openxlsx_deleted")
        }
        attr(.self$tables, "tableName") <- tableNames
        attr(.self$tables, "sheet") <- tableSheets
      }


      ## drawing will always be the first relationship and printerSettings second
      if (nSheets > 1) {
        for (i in 1:(nSheets - 1L)) {
          .self$worksheets_rels[[i]][1:3] <- genBaseSheetRels(i)
        }
      } else {
        .self$worksheets_rels <- list()
      }


      ## remove sheet
      sn <-
        unlist(lapply(workbook$sheets, function(x) {
          regmatches(
            x, regexpr('(?<= name=")[^"]+', x, perl = TRUE)
          )
        }))
      .self$workbook$sheets <- workbook$sheets[!sn %in% sheetName]

      ## Reset rIds
      if (nSheets > 1) {
        for (i in (sheet + 1L):nSheets) {
          .self$workbook$sheets <-
            gsub(stri_join("rId", i),
              stri_join("rId", i - 1L),
              workbook$sheets,
              fixed = TRUE
            )
        }
      } else {
        .self$workbook$sheets <- NULL
      }

      ## Can remove highest sheet
      .self$workbook.xml.rels <-
        workbook.xml.rels[!grepl(sprintf("sheet%s.xml", nSheets), workbook.xml.rels)]

      ## definedNames
      if (length(workbook$definedNames) > 0) {
        belongTo <- getDefinedNamesSheet(workbook$definedNames)
        .self$workbook$definedNames <-
          workbook$definedNames[!belongTo %in% sheetName]
      }

      invisible(1)
    },

    addDXFS = function(style) {
      dxf <- "<dxf>"
      dxf <- stri_join(dxf, createFontNode(style))
      fillNode <- NULL

      if (!is.null(style$fill$fillFg) | !is.null(style$fill$fillBg)) {
        dxf <- stri_join(dxf, createFillNode(style))
      }

      if (any(!is.null(
        c(
          style$borderLeft,
          style$borderRight,
          style$borderTop,
          style$borderBottom,
          style$borderDiagonal
        )
      ))) {
        dxf <- stri_join(dxf, createBorderNode(style))
      }

      dxf <- stri_join(dxf, "</dxf>", sep = " ")
      if (dxf %in% styles$dxfs) {
        return(which(styles$dxfs == dxf) - 1L)
      }

      dxfId <- length(styles$dxfs)
      .self$styles$dxfs <- c(styles$dxfs, dxf)

      return(dxfId)
    },

    dataValidation = function(sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      type,
      operator,
      value,
      allowBlank,
      showInputMsg,
      showErrorMsg) {
      sheet <- validateSheet(sheet)
      sqref <-
        stri_join(getCellRefs(data.frame(
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


      if (type == "date") {
        origin <- 25569L
        if (grepl(
          'date1904="1"|date1904="true"',
          stri_join(unlist(workbook), sep = " ", collapse = ""),
          ignore.case = TRUE
        )) {
          origin <- 24107L
        }

        value <- as.integer(value) + origin
      }

      if (type == "time") {
        origin <- 25569L
        if (grepl(
          'date1904="1"|date1904="true"',
          stri_join(unlist(workbook), sep = " ", collapse = ""),
          ignore.case = TRUE
        )) {
          origin <- 24107L
        }

        t <- format(value[1], "%z")
        offSet <-
          suppressWarnings(ifelse(substr(t, 1, 1) == "+", 1L, -1L) * (as.integer(substr(t, 2, 3)) + as.integer(substr(t, 4, 5)) / 60) / 24)
        if (is.na(offSet)) {
          offSet[i] <- 0
        }

        value <- as.numeric(as.POSIXct(value)) / 86400 + origin + offSet
      }

      form <-
        sapply(seq_along(value), function(i) {
          sprintf("<formula%s>%s</formula%s>", i, value[i], i)
        })
      .self$worksheets[[sheet]]$dataValidations <-
        c(
          worksheets[[sheet]]$dataValidations,
          stri_join(header, stri_join(form, collapse = ""), "</dataValidation>")
        )

      invisible(0)
    },

    dataValidation_list = function(sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      value,
      allowBlank,
      showInputMsg,
      showErrorMsg) {
      sheet <- validateSheet(sheet)
      sqref <-
        stri_join(getCellRefs(data.frame(
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
          showErrorMsg,
          sqref
        )

      formula <-
        sprintf("<x14:formula1><xm:f>%s</xm:f></x14:formula1>", value)
      sqref <- sprintf("<xm:sqref>%s</xm:sqref>", sqref)

      xmlData <-
        stri_join(data_val, formula, sqref, "</x14:dataValidation>")

      .self$worksheets[[sheet]]$dataValidationsLst <-
        c(worksheets[[sheet]]$dataValidationsLst, xmlData)

      invisible(0)
    },

    conditionalFormatting = function(sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      dxfId,
      formula,
      type,
      values,
      params) {
      sheet <- validateSheet(sheet)
      sqref <-
        stri_join(getCellRefs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        )), collapse = ":")



      ## Increment priority of conditional formatting rule
      if (length(worksheets[[sheet]]$conditionalFormatting) > 0) {
        for (i in length(worksheets[[sheet]]$conditionalFormatting):1) {
          priority <-
            regmatches(
              worksheets[[sheet]]$conditionalFormatting[[i]],
              regexpr(
                '(?<=priority=")[0-9]+',
                worksheets[[sheet]]$conditionalFormatting[[i]],
                perl = TRUE
              )
            )
          priority_new <- as.integer(priority) + 1L

          priority_pattern <- sprintf('priority="%s"', priority)
          priority_new <- sprintf('priority="%s"', priority_new)

          ## now replace
          .self$worksheets[[sheet]]$conditionalFormatting[[i]] <-
            gsub(priority_pattern,
              priority_new,
              worksheets[[sheet]]$conditionalFormatting[[i]],
              fixed = TRUE
            )
        }
      }

      nms <- c(names(worksheets[[sheet]]$conditionalFormatting), sqref)

      if (type == "colorScale") {
        ## formula contains the colours
        ## values contains numerics or is NULL
        ## dxfId is ignored

        if (is.null(values)) {
          if (length(formula) == 2L) {
            cfRule <-
              sprintf(
                '<cfRule type="colorScale" priority="1"><colorScale>
                             <cfvo type="min"/><cfvo type="max"/>
                             <color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>',
                formula[[1]],
                formula[[2]]
              )
          } else {
            cfRule <-
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
            cfRule <-
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
            cfRule <-
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
      } else if (type == "dataBar") {
        # forumula is a vector of colours of length 1 or 2
        # values is NULL or a numeric vector of equal length as formula

        if (length(formula) == 2L) {
          negColour <- formula[[1]]
          posColour <- formula[[2]]
        } else {
          posColour <- formula
          negColour <- "FFFF0000"
        }

        guid <-
          stri_join(
            "F7189283-14F7-4DE0-9601-54DE9DB",
            40000L + length(worksheets[[sheet]]$extLst)
          )

        showValue <- 1
        if ("showValue" %in% names(params)) {
          showValue <- as.integer(params$showValue)
        }

        gradient <- 1
        if ("gradient" %in% names(params)) {
          gradient <- as.integer(params$gradient)
        }

        border <- 1
        if ("border" %in% names(params)) {
          border <- as.integer(params$border)
        }

        if (is.null(values)) {
          cfRule <-
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
          cfRule <-
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

        .self$worksheets[[sheet]]$extLst <-
          c(
            worksheets[[sheet]]$extLst,
            gen_databar_extlst(
              guid = guid,
              sqref = sqref,
              posColour = posColour,
              negColour = negColour,
              values = values,
              border = border,
              gradient = gradient
            )
          )
      } else if (type == "expression") {
        cfRule <-
          sprintf(
            '<cfRule type="expression" dxfId="%s" priority="1"><formula>%s</formula></cfRule>',
            dxfId,
            formula
          )
      } else if (type == "duplicatedValues") {
        cfRule <-
          sprintf(
            '<cfRule type="duplicateValues" dxfId="%s" priority="1"/>',
            dxfId
          )
      } else if (type == "containsText") {
        cfRule <-
          sprintf(
            '<cfRule type="containsText" dxfId="%s" priority="1" operator="containsText" text="%s">
                        	<formula>NOT(ISERROR(SEARCH("%s", %s)))</formula>
                       </cfRule>',
            dxfId,
            values,
            values,
            unlist(strsplit(sqref, split = ":"))[[1]]
          )
      } else if (type == "notContainsText") {
        cfRule <-
          sprintf(
            '<cfRule type="notContainsText" dxfId="%s" priority="1" operator="notContains" text="%s">
                        	<formula>ISERROR(SEARCH("%s", %s))</formula>
                       </cfRule>',
            dxfId,
            values,
            values,
            unlist(strsplit(sqref, split = ":"))[[1]]
          )
      } else if (type == "beginsWith") {
        cfRule <-
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
      } else if (type == "endsWith") {
        cfRule <-
          sprintf(
            '<cfRule type="endsWith" dxfId="%s" priority="1" operator="endsWith" text="%s">
                        	<formula>RIGHT(%s,LEN("%s"))="%s"</formula>
                       </cfRule>',
            dxfId,
            values,

            unlist(strsplit(sqref, split = ":"))[[1]],
            values,
            values
          )
      } else if (type == "between") {
        cfRule <-
          sprintf(
            '<cfRule type="cellIs" dxfId="%s" priority="1" operator="between"><formula>%s</formula><formula>%s</formula></cfRule>',
            dxfId,
            formula[1],
            formula[2]
          )
      } else if (type == "topN") {
        cfRule <-
          sprintf(
            '<cfRule type="top10" dxfId="%s" priority="1" rank="%s" percent="%s"></cfRule>',
            dxfId,
            values[1],
            values[2]
          )
      } else if (type == "bottomN") {
        cfRule <-
          sprintf(
            '<cfRule type="top10" dxfId="%s" priority="1" rank="%s" percent="%s" bottom="1"></cfRule>',
            dxfId,
            values[1],
            values[2]
          )
      }

      .self$worksheets[[sheet]]$conditionalFormatting <-
        append(worksheets[[sheet]]$conditionalFormatting, cfRule)

      names(.self$worksheets[[sheet]]$conditionalFormatting) <- nms

      invisible(0)
    },

    mergeCells = function(sheet, startRow, endRow, startCol, endCol) {
      sheet <- validateSheet(sheetName = sheet)

      sqref <-
        getCellRefs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        ))
      exMerges <-
        regmatches(
          worksheets[[sheet]]$mergeCells,
          regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells)
        )

      if (!is.null(exMerges)) {
        comps <-
          lapply(exMerges, function(rectCoords) {
            unlist(strsplit(rectCoords, split = ":"))
          })
        exMergedCells <- build_cell_merges(comps = comps)
        newMerge <- unlist(build_cell_merges(comps = list(sqref)))

        ## Error if merge intersects
        mergeIntersections <-
          sapply(exMergedCells, function(x) {
            any(x %in% newMerge)
          })
        if (any(mergeIntersections)) {
          stop(
            sprintf(
              "Merge intersects with existing merged cells: \n\t\t%s.\nRemove existing merge first.",
              stri_join(exMerges[mergeIntersections], collapse = "\n\t\t")
            )
          )
        }
      }

      .self$worksheets[[sheet]]$mergeCells <-
        c(
          worksheets[[sheet]]$mergeCells,
          sprintf(
            '<mergeCell ref="%s"/>',
            stri_join(sqref,
              collapse = ":", sep =
                " "
            )
          )
        )
    },

    removeCellMerge = function(sheet, startRow, endRow, startCol, endCol) {
      sheet <- validateSheet(sheet)

      sqref <-
        getCellRefs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        ))
      exMerges <-
        regmatches(
          worksheets[[sheet]]$mergeCells,
          regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells)
        )

      if (!is.null(exMerges)) {
        comps <-
          lapply(exMerges, function(x) {
            unlist(strsplit(x, split = ":"))
          })
        exMergedCells <- build_cell_merges(comps = comps)
        newMerge <- unlist(build_cell_merges(comps = list(sqref)))

        ## Error if merge intersects
        mergeIntersections <-
          sapply(exMergedCells, function(x) {
            any(x %in% newMerge)
          })
      }

      ## Remove intersection
      .self$worksheets[[sheet]]$mergeCells <-
        worksheets[[sheet]]$mergeCells[!mergeIntersections]
    },

    freezePanes = function(sheet,
      firstActiveRow = NULL,
      firstActiveCol = NULL,
      firstRow = FALSE,
      firstCol = FALSE) {
      sheet <- validateSheet(sheet)
      paneNode <- NULL

      if (firstRow) {
        paneNode <-
          '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
      } else if (firstCol) {
        paneNode <-
          '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>'
      }


      if (is.null(paneNode)) {
        if (firstActiveRow == 1 & firstActiveCol == 1) {
          ## nothing to do
          return(NULL)
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
          attrs <-
            sprintf(
              'ySplit="%s" xSplit="%s"',
              firstActiveRow - 1L,
              firstActiveCol - 1L
            )
          activePane <- "bottomRight"
        }

        topLeftCell <-
          getCellRefs(data.frame(firstActiveRow, firstActiveCol))

        paneNode <-
          sprintf(
            '<pane %s topLeftCell="%s" activePane="%s" state="frozen"/><selection pane="%s"/>',
            stri_join(attrs, collapse = " ", sep = " "),
            topLeftCell,
            activePane,
            activePane
          )
      }

      .self$worksheets[[sheet]]$freezePane <- paneNode
    },

    insertImage = function(sheet,
      file,
      startRow,
      startCol,
      width,
      height,
      rowOffset = 0,
      colOffset = 0) {
      ## within the sheet the drawing node's Id refernce an id in the sheetRels
      ## sheet rels reference the drawingi.xml file
      ## drawingi.xml refernece drawingRels
      ## drawing rels reference an image in the media folder
      ## worksheetRels(sheet(i)) references drawings(j)

      sheet <- validateSheet(sheet)

      imageType <- regmatches(file, gregexpr("\\.[a-zA-Z]*$", file))
      imageType <- gsub("^\\.", "", imageType)

      imageNo <- length((drawings[[sheet]])) + 1L
      mediaNo <- length(media) + 1L

      startCol <- convertFromExcelRef(startCol)

      ## update Content_Types
      if (!any(grepl(stri_join("image/", imageType), Content_Types))) {
        .self$Content_Types <-
          unique(c(
            sprintf(
              '<Default Extension="%s" ContentType="image/%s"/>',
              imageType,
              imageType
            ),
            Content_Types
          ))
      }

      ## drawings rels (Reference from drawings.xml to image file in media folder)
      .self$drawings_rels[[sheet]] <- c(
        drawings_rels[[sheet]],
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
      .self$media <- append(media, tmp)

      ## create drawing.xml
      anchor <-
        '<xdr:oneCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'

      from <- sprintf(
        '<xdr:from xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
    <xdr:col xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">%s</xdr:col>
    <xdr:colOff xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">%s</xdr:colOff>
    <xdr:row xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">%s</xdr:row>
    <xdr:rowOff xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">%s</xdr:rowOff>
  </xdr:from>',
        startCol - 1L,
        colOffset,
        startRow - 1L,
        rowOffset
      )

      drawingsXML <- stri_join(
        anchor,
        from,
        sprintf(
          '<xdr:ext xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" cx="%s" cy="%s"/>',
          width,
          height
        ),
        genBasePic(imageNo),
        "<xdr:clientData/>",
        "</xdr:oneCellAnchor>"
      )


      ## append to workbook drawing
      .self$drawings[[sheet]] <- c(drawings[[sheet]], drawingsXML)
    },

    preSaveCleanUp = function() {
      ## Steps
      # Order workbook.xml.rels:
      #   sheets -> style -> theme -> sharedStrings -> persons -> tables -> calcChain
      # Assign workbook.xml.rels children rIds, seq_along(workbook.xml.rels)
      # Assign workbook$sheets rIds 1:nSheets
      #
      ## drawings will always be r:id1 on worksheet
      ## tables will always have r:id equal to table xml file number tables/table(i).xml

      ## Every worksheet has a drawingXML as r:id 1
      ## Every worksheet has a printerSettings as r:id 2
      ## Tables from r:id 3 to nTables+3 - 1
      ## HyperLinks from nTables+3 to nTables+3+nHyperLinks-1
      ## vmlDrawing to have rId

      sheetRIds <-
        as.integer(unlist(regmatches(
          workbook$sheets,
          gregexpr('(?<=r:id="rId)[0-9]+', workbook$sheets, perl = TRUE)
        )))

      nSheets <- length(sheetRIds)
      nExtRefs <- length(externalLinks)
      nPivots <- length(pivotDefinitions)

      ## add a worksheet if none added
      if (nSheets == 0) {
        warning("Workbook does not contain any worksheets. A worksheet will be added.",
          call. = FALSE
        )
        .self$addWorksheet("Sheet 1")
        nSheets <- 1L
      }

      ## get index of each child element for ordering
      sheetInds <-
        which(grepl(
          "(worksheets|chartsheets)/sheet[0-9]+\\.xml",
          workbook.xml.rels
        ))
      stylesInd <- which(grepl("styles\\.xml", workbook.xml.rels))
      themeInd <-
        which(grepl("theme/theme[0-9]+.xml", workbook.xml.rels))
      connectionsInd <-
        which(grepl("connections.xml", workbook.xml.rels))
      extRefInds <-
        which(grepl("externalLinks/externalLink[0-9]+.xml", workbook.xml.rels))
      sharedStringsInd <-
        which(grepl("sharedStrings.xml", workbook.xml.rels))
      tableInds <- which(grepl("table[0-9]+.xml", workbook.xml.rels))
      personInds <- which(grepl("person.xml", workbook.xml.rels))


      ## Reordering of workbook.xml.rels
      ## don't want to re-assign rIds for pivot tables or slicer caches
      pivotNode <-
        workbook.xml.rels[grepl(
          "pivotCache/pivotCacheDefinition[0-9].xml",
          workbook.xml.rels
        )]
      slicerNode <-
        workbook.xml.rels[which(grepl("slicerCache[0-9]+.xml", workbook.xml.rels))]

      ## Reorder children of workbook.xml.rels
      .self$workbook.xml.rels <-
        workbook.xml.rels[c(
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
      .self$workbook.xml.rels <-
        unlist(lapply(seq_along(workbook.xml.rels), function(i) {
          gsub('(?<=Relationship Id="rId)[0-9]+',
            i,
            workbook.xml.rels[[i]],
            perl = TRUE
          )
        }))

      .self$workbook.xml.rels <- c(workbook.xml.rels, pivotNode, slicerNode)



      if (!is.null(vbaProject)) {
        .self$workbook.xml.rels <-
          c(
            workbook.xml.rels,
            sprintf(
              '<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>',
              1L + length(workbook.xml.rels)
            )
          )
      }

      ## Reassign rId to workbook sheet elements, (order sheets by sheetId first)
      .self$workbook$sheets <-
        unlist(lapply(seq_along(workbook$sheets), function(i) {
          gsub('(?<= r:id="rId)[0-9]+', i, workbook$sheets[[i]], perl = TRUE)
        }))

      ## re-order worksheets if need to
      if (any(sheetOrder != 1:nSheets)) {
        .self$workbook$sheets <- workbook$sheets[sheetOrder]
      }



      ## re-assign tabSelected
      state <- rep.int("visible", nSheets)
      state[grepl("hidden", workbook$sheets)] <- "hidden"
      visible_sheet_index <- which(state %in% "visible")[[1]]

      .self$workbook$bookViews <-
        sprintf(
          '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="13125" windowHeight="6105" firstSheet="%s" activeTab="%s"/></bookViews>',
          visible_sheet_index - 1L,
          visible_sheet_index - 1L
        )

      .self$worksheets[[visible_sheet_index]]$sheetViews <-
        sub(
          '( tabSelected="0")|( tabSelected="false")',
          ' tabSelected="1"',
          worksheets[[visible_sheet_index]]$sheetViews,
          ignore.case = TRUE
        )
      if (nSheets > 1) {
        for (i in (1:nSheets)[!(1:nSheets) %in% visible_sheet_index]) {
          .self$worksheets[[i]]$sheetViews <-
            sub(
              ' tabSelected="(1|true|false|0)"',
              ' tabSelected="0"',
              worksheets[[i]]$sheetViews,
              ignore.case = TRUE
            )
        }
      }





      if (length(workbook$definedNames) > 0) {
        sheetNames <- sheet_names[sheetOrder]

        belongTo <- getDefinedNamesSheet(workbook$definedNames)

        ## sheetNames is in re-ordered order (order it will be displayed)
        newId <- match(belongTo, sheetNames) - 1L
        oldId <-
          as.numeric(regmatches(
            workbook$definedNames,
            regexpr(
              '(?<= localSheetId=")[0-9]+',
              workbook$definedNames,
              perl = TRUE
            )
          ))

        for (i in seq_along(workbook$definedNames)) {
          if (!is.na(newId[i])) {
            .self$workbook$definedNames[[i]] <-
              gsub(
                sprintf('localSheetId=\"%s\"', oldId[i]),
                sprintf('localSheetId=\"%s\"', newId[i]),
                workbook$definedNames[[i]],
                fixed = TRUE
              )
          }
        }
      }




      ## update workbook r:id to match reordered workbook.xml.rels externalLink element
      if (length(extRefInds) > 0) {
        newInds <- as.integer(seq_along(extRefInds) + length(sheetInds))
        .self$workbook$externalReferences <-
          stri_join(
            "<externalReferences>",
            stri_join(
              sprintf('<externalReference r:id=\"rId%s\"/>', newInds),
              collapse = ""
            ),
            "</externalReferences>"
          )
      }

      ## styles
      numFmtIds <- 50000L
      #for (i in which(!isChartSheet)) {
      #  worksheets[[i]]$sheet_data$style_id <-
      #    rep.int(x = as.integer(NA), times = worksheets[[i]]$sheet_data$n_elements)
      #}


      for (x in styleObjects) {
        if (length(x$rows) > 0 & length(x$cols) > 0) {
          this.sty <- x$style$copy()

          if (!is.null(this.sty$numFmt)) {
            if (this.sty$numFmt$numFmtId == 9999) {
              this.sty$numFmt$numFmtId <- numFmtIds
              numFmtIds <- numFmtIds + 1L
            }
          }


          ## convert sheet name to index
          sheet <- which(sheet_names == x$sheet)
          sId <-
            .self$updateStyles(this.sty) ## this creates the XML for styles.XML

          cells_to_style <- stri_join(x$rows, x$cols, sep = ",")
          existing_cells <-
            stri_join(worksheets[[sheet]]$sheet_data$rows,
              worksheets[[sheet]]$sheet_data$cols,
              sep = ","
            )

          ## In here we create any style_ids that don't yet exist in sheet_data
          #worksheets[[sheet]]$sheet_data$style_id[existing_cells %in% cells_to_style] <-
          #  sId


          new_cells_to_append <-
            which(!cells_to_style %in% existing_cells)
          if (length(new_cells_to_append) > 0) {
            #worksheets[[sheet]]$sheet_data$style_id <-
            #  c(
            #    worksheets[[sheet]]$sheet_data$style_id,
            #    rep.int(x = sId, times = length(new_cells_to_append))
            #  )

            .self$worksheets[[sheet]]$sheet_data$rows <-
              c(worksheets[[sheet]]$sheet_data$rows, x$rows[new_cells_to_append])
            .self$worksheets[[sheet]]$sheet_data$cols <-
              c(worksheets[[sheet]]$sheet_data$cols, x$cols[new_cells_to_append])
            .self$worksheets[[sheet]]$sheet_data$t <-
              c(worksheets[[sheet]]$sheet_data$t, rep(as.integer(NA), length(new_cells_to_append)))
            .self$worksheets[[sheet]]$sheet_data$v <-
              c(
                worksheets[[sheet]]$sheet_data$v,
                rep(as.character(NA), length(new_cells_to_append))
              )
            .self$worksheets[[sheet]]$sheet_data$f <-
              c(
                worksheets[[sheet]]$sheet_data$f,
                rep(as.character(NA), length(new_cells_to_append))
              )
            .self$worksheets[[sheet]]$sheet_data$data_count <-
              worksheets[[sheet]]$sheet_data$data_count + 1L

            .self$worksheets[[sheet]]$sheet_data$n_elements <-
              as.integer(length(worksheets[[sheet]]$sheet_data$rows))
          }
        }
      }


      ## Make sure all rowHeights have rows, if not append them!
      for (i in seq_along(worksheets)) {
        if (length(rowHeights[[i]]) > 0) {
          rh <- as.integer(names(rowHeights[[i]]))
          missing_rows <- rh[!rh %in% worksheets[[i]]$sheet_data$rows]
          n <- length(missing_rows)

          if (n > 0) {
            #worksheets[[i]]$sheet_data$style_id <-
            #  c(
            #    worksheets[[i]]$sheet_data$style_id,
            #    rep.int(as.integer(NA), times = n)
            #  )

            .self$worksheets[[i]]$sheet_data$rows <-
              c(worksheets[[i]]$sheet_data$rows, missing_rows)
            .self$worksheets[[i]]$sheet_data$cols <-
              c(
                worksheets[[i]]$sheet_data$cols,
                rep.int(as.integer(NA), times = n)
              )

            .self$worksheets[[i]]$sheet_data$t <-
              c(worksheets[[i]]$sheet_data$t, rep(as.integer(NA), times = n))
            .self$worksheets[[i]]$sheet_data$v <-
              c(
                worksheets[[i]]$sheet_data$v,
                rep(as.character(NA), times = n)
              )
            .self$worksheets[[i]]$sheet_data$f <-
              c(
                worksheets[[i]]$sheet_data$f,
                rep(as.character(NA), times = n)
              )
            .self$worksheets[[i]]$sheet_data$data_count <-
              worksheets[[i]]$sheet_data$data_count + 1L

            .self$worksheets[[i]]$sheet_data$n_elements <-
              as.integer(length(worksheets[[i]]$sheet_data$rows))
          }
        }

        ## write colwidth and coloutline XML
        if (length(colWidths[[i]]) > 0) {
          invisible(.self$setColWidths(i))
        }

        if (length(colOutlineLevels[[i]]) > 0) {
          invisible(.self$groupColumns(i))
        }
      }
    },

    addStyle = function(sheet, style, rows, cols, stack) {
      sheet <- sheet_names[[sheet]]

      if (length(styleObjects) == 0) {
        .self$styleObjects <- list(list(
          style = style,
          sheet = sheet,
          rows = rows,
          cols = cols
        ))
      } else if (stack) {
        nStyles <- length(styleObjects)

        ## ********** Assume all styleObjects cells have one a single worksheet **********
        ## Loop through existing styleObjects
        newInds <- 1:length(rows)
        keepStyle <- rep(TRUE, nStyles)
        for (i in 1:nStyles) {
          if (sheet == styleObjects[[i]]$sheet) {
            ## Now check rows and cols intersect
            ## toRemove are the elements that the new style doesn't apply to, we remove these from the style object as it
            ## is copied, merged with the new style and given the new data points

            ex_row_cols <-
              stri_join(styleObjects[[i]]$rows, styleObjects[[i]]$cols, sep = "-")
            new_row_cols <- stri_join(rows, cols, sep = "-")


            ## mergeInds are the intersection of the two styles that will need to merge
            mergeInds <- which(new_row_cols %in% ex_row_cols)

            ## newInds are inds that don't exist in the current - this cumulates until the end to see if any are new
            newInds <- newInds[!newInds %in% mergeInds]


            ## If the new style does not merge
            if (length(mergeInds) > 0) {
              to_remove_from_this_style_object <-
                which(ex_row_cols %in% new_row_cols)

              ## the new style intersects with this styleObjects[[i]], we need to remove the intersecting rows and
              ## columns from styleObjects[[i]]
              if (length(to_remove_from_this_style_object) > 0) {
                ## remove these from style object
                .self$styleObjects[[i]]$rows <-
                  styleObjects[[i]]$rows[-to_remove_from_this_style_object]
                .self$styleObjects[[i]]$cols <-
                  styleObjects[[i]]$cols[-to_remove_from_this_style_object]

                if (length(styleObjects[[i]]$rows) == 0 |
                    length(styleObjects[[i]]$cols) == 0) {
                  keepStyle[i] <-
                    FALSE
                } ## this style applies to no rows or columns anymore
              }

              ## append style object for intersecting cells

              ## we are appending a new style
              keepStyle <-
                c(keepStyle, TRUE) ## keepStyle is used to remove styles that apply to 0 rows OR 0 columns

              ## Merge Style and append to styleObjects
              .self$styleObjects <-
                append(styleObjects, list(
                  list(
                    style = mergeStyle(styleObjects[[i]]$style, newStyle = style),
                    sheet = sheet,
                    rows = rows[mergeInds],
                    cols = cols[mergeInds]
                  )
                ))
            }
          } ## if sheet == styleObjects[[i]]$sheet
        } ## End of loop through styles

        ## remove any styles that no longer have any affect
        if (!all(keepStyle)) {
          .self$styleObjects <- styleObjects[keepStyle]
        }

        ## append style object for non-intersecting cells
        if (length(newInds) > 0) {
          .self$styleObjects <- append(styleObjects, list(list(
            style = style,
            sheet = sheet,
            rows = rows[newInds],
            cols = cols[newInds]
          )))
        }
      } else {
        ## else we are not stacking

        .self$styleObjects <- append(styleObjects, list(list(
          style = style,
          sheet = sheet,
          rows = rows,
          cols = cols
        )))
      } ## End if(length(styleObjects) > 0) else if(stack) {}
    },

    createNamedRegion = function(ref1, ref2, name, sheet, localSheetId = NULL) {
      name <- replaceIllegalCharacters(name)

      if (is.null(localSheetId)) {
        .self$workbook$definedNames <- c(
          workbook$definedNames,
          sprintf(
            '<definedName name="%s">\'%s\'!%s:%s</definedName>',
            name,
            sheet,
            ref1,
            ref2
          )
        )
      } else {
        .self$workbook$definedNames <- c(
          workbook$definedNames,
          sprintf(
            '<definedName name="%s" localSheetId="%s">\'%s\'!%s:%s</definedName>',
            name,
            localSheetId,
            sheet,
            ref1,
            ref2
          )
        )
      }
    },

    validate_table_name = function(tableName) {
      tableName <-
        tolower(tableName) ## Excel forces named regions to lowercase

      if (nchar(tableName) > 255) {
        stop("tableName must be less than 255 characters.")
      }

      if (grepl("$", tableName, fixed = TRUE)) {
        stop("'$' character cannot exist in a tableName")
      }

      if (grepl(" ", tableName, fixed = TRUE)) {
        stop("spaces cannot exist in a table name")
      }

      # if(!grepl("^[A-Za-z_]", tableName, perl = TRUE))
      #   stop("tableName must begin with a letter or an underscore")

      if (grepl("R[0-9]+C[0-9]+",
        tableName,
        perl = TRUE,
        ignore.case = TRUE
      )) {
        stop("tableName cannot be the same as a cell reference, such as R1C1")
      }

      if (grepl("^[A-Z]{1,3}[0-9]+$", tableName, ignore.case = TRUE)) {
        stop("tableName cannot be the same as a cell reference")
      }

      if (tableName %in% attr(tables, "tableName")) {
        stop(sprintf("Table with name '%s' already exists!", tableName))
      }

      return(tableName)
    },

    check_overwrite_tables = function(sheet,
      new_rows,
      new_cols,
      error_msg = "Cannot overwrite existing table with another table.",
      check_table_header_only = FALSE) {
      ## check not overwriting another table
      if (length(tables) > 0) {
        tableSheets <- attr(tables, "sheet")
        sheetNo <- validateSheet(sheet)

        to_check <-
          which(tableSheets %in% sheetNo &
              !grepl("openxlsx_deleted", attr(tables, "tableName"), fixed = TRUE))

        if (length(to_check) > 0) {
          ## only look at tables on this sheet

          exTable <- tables[to_check]

          rows <-
            lapply(names(exTable), function(rectCoords) {
              as.numeric(unlist(regmatches(
                rectCoords, gregexpr("[0-9]+", rectCoords)
              )))
            })
          cols <-
            lapply(names(exTable), function(rectCoords) {
              convertFromExcelRef(unlist(regmatches(
                rectCoords, gregexpr("[A-Z]+", rectCoords)
              )))
            })

          if (check_table_header_only) {
            rows <- lapply(rows, function(x) {
              c(x[1], x[1])
            })
          }


          ## loop through existing tables checking if any over lap with new table
          for (i in 1:length(exTable)) {
            existing_cols <- cols[[i]]
            existing_rows <- rows[[i]]

            if ((min(new_cols) <= max(existing_cols)) &
                (max(new_cols) >= min(existing_cols)) &
                (min(new_rows) <= max(existing_rows)) &
                (max(new_rows) >= min(existing_rows))) {
              stop(error_msg)
            }
          }
        } ## end if(sheet %in% tableSheets)
      } ## end (length(tables) > 0)

      invisible(0)
    },

    show = function() {
      exSheets <- sheet_names
      nSheets <- length(exSheets)
      nImages <- length(media)
      nCharts <- length(charts)
      nStyles <- length(styleObjects)

      exSheets <- replaceXMLEntities(exSheets)
      showText <- "A Workbook object.\n"

      ## worksheets
      if (nSheets > 0) {
        showText <- c(showText, "\nWorksheets:\n")

        sheetTxt <- lapply(1:nSheets, function(i) {
          tmpTxt <- sprintf('Sheet %s: "%s"\n', i, exSheets[[i]])

          if (length(rowHeights[[i]]) > 0) {
            tmpTxt <-
              append(
                tmpTxt,
                c(
                  "\n\tCustom row heights (row: height)\n\t",
                  stri_join(
                    sprintf("%s: %s", names(rowHeights[[i]]), round(as.numeric(
                      rowHeights[[i]]
                    ), 2)),
                    collapse = ", ",
                    sep = " "
                  )
                )
              )
          }

          if (length(outlineLevels[[i]]) > 0) {
            tmpTxt <-
              append(
                tmpTxt,
                c(
                  "\n\tGrouped rows:\n\t",
                  stri_join(
                    sprintf("%s", names(outlineLevels[[i]])),
                    collapse = ", ",
                    sep = " "
                  )
                )
              )
          }

          if (length(colOutlineLevels[[i]]) > 0) {
            tmpTxt <-
              append(
                tmpTxt,
                c(
                  "\n\tGrouped columns:\n\t",
                  stri_join(
                    sprintf("%s", names(colOutlineLevels[[i]])),
                    collapse = ", ",
                    sep = " "
                  )
                )
              )
          }

          if (length(colWidths[[i]]) > 0) {
            cols <- names(colWidths[[i]])
            widths <- unname(colWidths[[i]])

            widths[widths != "auto"] <-
              as.numeric(widths[widths != "auto"])
            tmpTxt <-
              append(
                tmpTxt,
                c(
                  "\n\tCustom column widths (column: width)\n\t ",
                  stri_join(
                    sprintf("%s: %s", cols, substr(widths, 1, 5)),
                    sep = " ",
                    collapse = ", "
                  )
                )
              )
            tmpTxt <- c(tmpTxt, "\n")
          }
          c(tmpTxt, "\n\n")
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
            sprintf('Image %s: "%s"\n', 1:nImages, media)
          )
      }

      if (nCharts > 0) {
        showText <-
          c(
            showText,
            "\nCharts:\n",
            sprintf('Chart %s: "%s"\n', 1:nCharts, charts)
          )
      }

      if (nSheets > 0) {
        showText <-
          c(showText, sprintf(
            "Worksheet write order: %s",
            stri_join(sheetOrder, sep = " ", collapse = ", ")
          ))
      }

      cat(unlist(showText))
    },


    conditionalFormatCell = function(sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      dxfId,
      formula,
      type) {
      .Deprecated()

      sheet <- validateSheet(sheet)
      sqref <-
        stri_join(getCellRefs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        )), collapse = ":")

      ## Increment priority of conditional formatting rule
      if (length((worksheets[[sheet]]$conditionalFormatting)) > 0) {
        for (i in length(worksheets[[sheet]]$conditionalFormatting):1) {
          .self$worksheets[[sheet]]$conditionalFormatting[[i]] <-
            gsub('(?<=priority=")[0-9]+',
              i + 1L,
              worksheets[[sheet]]$conditionalFormatting[[i]],
              perl = TRUE
            )
        }
      }

      nms <- c(names(worksheets[[sheet]]$conditionalFormatting), sqref)

      if (type == "expression") {
        cfRule <-
          sprintf(
            '<cfRule type="expression" dxfId="%s" priority="1"><formula>%s</formula></cfRule>',
            dxfId,
            formula
          )
      } else if (type == "dataBar") {
        if (length(formula) == 2) {
          negColour <- formula[[1]]
          posColour <- formula[[2]]
        } else {
          posColour <- formula
          negColour <- "FFFF0000"
        }

        guid <-
          stri_join(
            "F7189283-14F7-4DE0-9601-54DE9DB",
            40000L + length(worksheets[[sheet]]$extLst)
          )
        cfRule <-
          sprintf(
            '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="%s"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{%s}</x14:id></ext></extLst></cfRule>',
            posColour,
            guid
          )
      } else if (length(formula) == 2L) {
        cfRule <-
          sprintf(
            '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="%s"/><color rgb="%s"/></colorScale></cfRule>',
            formula[[1]],
            formula[[2]]
          )
      } else {
        cfRule <-
          sprintf(
            '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="%s"/><color rgb="%s"/><color rgb="%s"/></colorScale></cfRule>',
            formula[[1]],
            formula[[2]],
            formula[[3]]
          )
      }

      .self$worksheets[[sheet]]$conditionalFormatting <-
        append(worksheets[[sheet]]$conditionalFormatting, cfRule)

      names(.self$worksheets[[sheet]]$conditionalFormatting) <- nms

      invisible(0)
    },

    loadStyles = function(stylesXML) {
      ## Build style objects from the styles XML
      styles_XML <- read_xml(stylesXML)

      ## Indexed colours
      vals <- xml_node(styles_XML, "styleSheet", "colors", "indexedColors")
      if (length(vals) > 0) {
        .self$styles$indexedColors <-
          stri_join("<colors>", vals, "</colors>")
      }

      .self$styles$numFmts <- numFmts <- xml_node(styles_XML, "styleSheet", "numFmts", "numFmt")
      # numFmts_attr <- getXMLattr(numFmts, "numFmt")

      ## dxf
      .self$styles$dxfs <- dxf <- xml_node(styles_XML, "styleSheet", "dxfs", "dxf")

      tableStyles <- xml_node(styles_XML, "styleSheet", "tableStyles")
      if (length(tableStyles) > 0) {
        .self$styles$tableStyles <- tableStyles
      }

      # extLst <- getXML2(styles_XML, "styleSheet", "extLst")
      # if (length(extLst) > 0) {
      #  styles$extLst <- extLst
      # }

      .self$styles$borders <- borders <- xml_node(styles_XML, "styleSheet", "borders", "border")

      .self$styles$fills <- fills <- xml_node(styles_XML, "styleSheet", "fills", "fill")

      .self$styles$cellStyleXfs <- cellStyleXfs <- xml_node(styles_XML, "styleSheet", "cellStyleXfs", "xf")

      .self$styles$cellXfs <- cellXfs <- xml_node(styles_XML, "styleSheet", "cellXfs", "xf")

      ## Number formats
      # numFmts <- getChildlessNode(xml = stylesTxt, tag = "numFmt")
      numFmts <- xml_node(styles_XML, "styleSheet", "numFmts", "numFmt")
      numFmtFlag <- FALSE
      if (length(numFmts) > 0) {
        numFmtsIds <-
          sapply(numFmts, getAttr, tag = 'numFmtId="', USE.NAMES = FALSE)
        formatCodes <-
          sapply(numFmts, getAttr, tag = 'formatCode="', USE.NAMES = FALSE)
        numFmts <-
          lapply(1:length(numFmts), function(i) {
            list("numFmtId" = numFmtsIds[[i]], "formatCode" = formatCodes[[i]])
          })
        numFmtFlag <- TRUE
      }

      ## fonts will maintain, sz, color, name, family scheme

      .self$styles$fonts <- fonts <- xml_node(styles_XML, "styleSheet", "fonts", "font")
      fonts <- buildFontList(fonts)


      fills <- buildFillList(fills)

      borders <- sapply(borders, buildBorder, USE.NAMES = FALSE)


      ## ------------------------------ build styleObjects ------------------------------ ##

      xfVals <- getXMLattr(cellXfs, "xf")

      styleObjects_tmp <- list()
      flag <- FALSE
      for (s in xfVals) {
        style <- createStyle()
        if (any(s != "0")) {
          if ("fontId" %in% names(s)) {

            style$fontId <- as.integer(s[["fontId"]])

            if (s[["fontId"]] != "0") {
              thisFont <- fonts[[(as.integer(s[["fontId"]]) + 1)]]


              if ("sz" %in% names(thisFont)) {
                style$fontSize <- thisFont$sz
              }

              if ("name" %in% names(thisFont)) {
                style$fontName <- thisFont$name
              }

              if ("family" %in% names(thisFont)) {
                style$fontFamily <- thisFont$family
              }

              if ("color" %in% names(thisFont)) {
                style$fontColour <- thisFont$color
              }

              if ("scheme" %in% names(thisFont)) {
                style$fontScheme <- thisFont$scheme
              }

              flags <-
                c("bold", "italic", "underline") %in% names(thisFont)
              if (any(flags)) {
                style$fontDecoration <- NULL
                if (flags[[1]]) {
                  style$fontDecoration <-
                    append(style$fontDecoration, "BOLD")
                }

                if (flags[[2]]) {
                  style$fontDecoration <-
                    append(style$fontDecoration, "ITALIC")
                }

                if (flags[[3]]) {
                  style$fontDecoration <-
                    append(style$fontDecoration, "UNDERLINE")
                }
              }
            }
          }

          if ("numFmtId" %in% names(s)) {
            if (s[["numFmtId"]] != "0") {
              if (as.integer(s[["numFmtId"]]) < 164) {
                style$numFmt <- list(numFmtId = s[["numFmtId"]])
              } else if (numFmtFlag) {
                style$numFmt <- numFmts[[which(s[["numFmtId"]] == numFmtsIds)[1]]]
              }
            }
          }

          ## Border
          if ("borderId" %in% names(s)) {
            if (s[["borderId"]] != "0") {
              # & "applyBorder" %in% names(s)){

              border_ind <- as.integer(s[["borderId"]]) + 1L
              if (border_ind <= length(borders)) {
                thisBorder <- borders[[border_ind]]

                if ("borderLeft" %in% names(thisBorder)) {
                  style$borderLeft <- thisBorder$borderLeft
                  style$borderLeftColour <- thisBorder$borderLeftColour
                }

                if ("borderRight" %in% names(thisBorder)) {
                  style$borderRight <- thisBorder$borderRight
                  style$borderRightColour <-
                    thisBorder$borderRightColour
                }

                if ("borderTop" %in% names(thisBorder)) {
                  style$borderTop <- thisBorder$borderTop
                  style$borderTopColour <- thisBorder$borderTopColour
                }

                if ("borderBottom" %in% names(thisBorder)) {
                  style$borderBottom <- thisBorder$borderBottom
                  style$borderBottomColour <-
                    thisBorder$borderBottomColour
                }

                if ("borderDiagonal" %in% names(thisBorder)) {
                  style$borderDiagonal <- thisBorder$borderDiagonal
                  style$borderDiagonalColour <-
                    thisBorder$borderDiagonalColour
                }

                if ("borderDiagonalUp" %in% names(thisBorder)) {
                  style$borderDiagonalUp <-
                    thisBorder$borderDiagonalUp
                }

                if ("borderDiagonalDown" %in% names(thisBorder)) {
                  style$borderDiagonalDown <-
                    thisBorder$borderDiagonalDown
                }
              }
            }
          }

          ## alignment
          # applyAlignment <- "applyAlignment" %in% names(s)
          if ("horizontal" %in% names(s)) {
            # & applyAlignment)
            style$halign <- s[["horizontal"]]
          }

          if ("vertical" %in% names(s)) {
            style$valign <- s[["vertical"]]
          }

          if ("indent" %in% names(s)) {
            style$indent <- s[["indent"]]
          }

          if ("textRotation" %in% names(s)) {
            style$textRotation <- s[["textRotation"]]
          }

          ## wrap text
          if ("wrapText" %in% names(s)) {
            if (s[["wrapText"]] %in% c("1", "true")) {
              style$wrapText <- TRUE
            }
          }

          if ("fillId" %in% names(s)) {
            if (s[["fillId"]] != "0") {
              fillId <- as.integer(s[["fillId"]]) + 1L

              if ("fgColor" %in% names(fills[[fillId]])) {
                tmpFg <- fills[[fillId]]$fgColor
                tmpBg <- fills[[fillId]]$bgColor

                if (!is.null(tmpFg)) {
                  style$fill$fillFg <- tmpFg
                }

                if (!is.null(tmpFg)) {
                  style$fill$fillBg <- tmpBg
                }
              } else {
                style$fill <- fills[[fillId]]
              }
            }
          }


          if ("xfId" %in% names(s)) {
            if (s[["xfId"]] != "0") {
              style$xfId <- s[["xfId"]]
            }
          }
        } ## end if !all(s == "0")

        # Cell protection settings can be "0", so we cannot just skip all zeroes
        if ("locked" %in% names(s)) {
          style$locked <- (s[["locked"]] == "1")
        }

        if ("hidden" %in% names(s)) {
          style$hidden <- (s[["hidden"]] == "1")
        }

        ## we need to skip the first one as this is used as the base style
        if (flag) {
          styleObjects_tmp <- append(styleObjects_tmp, list(style))
        }

        flag <- TRUE
      } ## end of for loop through styles s in ...


      ## ------------------------------ build styleObjects Complete ------------------------------ ##


      return(styleObjects_tmp)
    },

    protectWorkbook = function(protect = TRUE,
      lockStructure = FALSE,
      lockWindows = FALSE,
      password = NULL) {
      attr <- c()
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
        .self$workbook$workbookProtection <-
          sprintf(
            "<workbookProtection %s/>",
            stri_join(
              names(attr),
              '="',
              attr,
              '"',
              collapse = " ",
              sep = ""
            )
          )
      } else {
        .self$workbook$workbookProtection <- ""
      }
    },

    addCreator = function(Creator = NULL) {
      if (!is.null(Creator)) {
        current_creator <-
          stri_match(core, regex = "<dc:creator>(.*?)</dc:creator>")[1, 2]
        .self$core <-
          stri_replace_all_fixed(
            core,
            pattern = current_creator,
            replacement = stri_c(current_creator, Creator, sep = ";")
          )
      }
    },

    getCreators = function() {
      current_creator <-
        stri_match(core, regex = "<dc:creator>(.*?)</dc:creator>")[1, 2]

      current_creator_vec <- as.character(stri_split_fixed(
        str = current_creator,
        pattern = ";",
        simplify = T
      ))

      return(current_creator_vec)
    },

    changeLastModifiedBy = function(LastModifiedBy = NULL) {
      if (!is.null(LastModifiedBy)) {
        current_LastModifiedBy <-
          stri_match(core, regex = "<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>")[1, 2]
        .self$core <-
          stri_replace_all_fixed(
            core,
            pattern = current_LastModifiedBy,
            replacement = LastModifiedBy
          )
      }
    },

    surroundingBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType) {
      sheet <- sheet_names[[validateSheet(sheet)]]
      ## steps
      # get column class
      # get corresponding base style

      for (i in 1:nCol) {
        tmp <- genBaseColStyle(colClasses[[i]])

        colStyle <- tmp$style
        specialFormat <- tmp$specialFormat

        ## create style objects
        sTop <- colStyle$copy()
        sMid <- colStyle$copy()
        sBot <- colStyle$copy()

        ## First column
        if (i == 1) {
          if (nRow == 1 & nCol == 1) {

            ## All
            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            sTop$borderBottom <- borderStyle
            sTop$borderBottomColour <- borderColour

            sTop$borderLeft <- borderStyle
            sTop$borderLeftColour <- borderColour

            sTop$borderRight <- borderStyle
            sTop$borderRightColour <- borderColour

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol
              )
            ))
          } else if (nCol == 1) {

            ## Top
            sTop$borderLeft <- borderStyle
            sTop$borderLeftColour <- borderColour

            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            sTop$borderRight <- borderStyle
            sTop$borderRightColour <- borderColour

            ## Middle
            sMid$borderLeft <- borderStyle
            sMid$borderLeftColour <- borderColour

            sMid$borderRight <- borderStyle
            sMid$borderRightColour <- borderColour

            ## Bottom
            sBot$borderBottom <- borderStyle
            sBot$borderBottomColour <- borderColour

            sBot$borderLeft <- borderStyle
            sBot$borderLeftColour <- borderColour

            sBot$borderRight <- borderStyle
            sBot$borderRightColour <- borderColour

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol
              )
            ))

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sMid,
                "sheet" = sheet,
                "rows" = (startRow + 1L):(startRow + nRow - 2L), # 2nd -> 2nd to last
                "cols" = rep.int(startCol, nRow - 2L)
              )
            ))

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sBot,
                "sheet" = sheet,
                "rows" = startRow + nRow - 1L,
                "cols" = startCol
              )
            ))
          } else if (nRow == 1) {

            ## All
            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            sTop$borderBottom <- borderStyle
            sTop$borderBottomColour <- borderColour

            sTop$borderLeft <- borderStyle
            sTop$borderLeftColour <- borderColour

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol
              )
            ))
          } else {

            ## Top
            sTop$borderLeft <- borderStyle
            sTop$borderLeftColour <- borderColour

            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            ## Middle
            sMid$borderLeft <- borderStyle
            sMid$borderLeftColour <- borderColour

            ## Bottom
            sBot$borderLeft <- borderStyle
            sBot$borderLeftColour <- borderColour

            sBot$borderBottom <- borderStyle
            sBot$borderBottomColour <- borderColour

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol
              )
            ))

            if (nRow > 2) {
              .self$styleObjects <- append(styleObjects, list(
                list(
                  "style" = sMid,
                  "sheet" = sheet,
                  "rows" = (startRow + 1L):(startRow + nRow - 2L), # 2nd -> 2nd to last
                  "cols" = rep.int(startCol, nRow - 2L)
                )
              ))
            }

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sBot,
                "sheet" = sheet,
                "rows" = startRow + nRow - 1L,
                "cols" = startCol
              )
            ))
          }
        } else if (i == nCol) {
          if (nRow == 1) {

            ## All
            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            sTop$borderBottom <- borderStyle
            sTop$borderBottomColour <- borderColour

            sTop$borderRight <- borderStyle
            sTop$borderRightColour <- borderColour

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol + nCol - 1L
              )
            ))
          } else {

            ## Top
            sTop$borderRight <- borderStyle
            sTop$borderRightColour <- borderColour

            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            ## Middle
            sMid$borderRight <- borderStyle
            sMid$borderRightColour <- borderColour

            ## Bottom
            sBot$borderRight <- borderStyle
            sBot$borderRightColour <- borderColour

            sBot$borderBottom <- borderStyle
            sBot$borderBottomColour <- borderColour

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol + nCol - 1L
              )
            ))

            if (nRow > 2) {
              .self$styleObjects <- append(styleObjects, list(
                list(
                  "style" = sMid,
                  "sheet" = sheet,
                  "rows" = (startRow + 1L):(startRow + nRow - 2L), # 2nd -> 2nd to last
                  "cols" = rep.int(startCol + nCol - 1L, nRow - 2L)
                )
              ))
            }


            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sBot,
                "sheet" = sheet,
                "rows" = startRow + nRow - 1L,
                "cols" = startCol + nCol - 1L
              )
            ))
          }
        } else { ## inside columns

          if (nRow == 1) {

            ## Top
            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            ## Bottom
            sTop$borderBottom <- borderStyle
            sTop$borderBottomColour <- borderColour

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol + i - 1L
              )
            ))
          } else {

            ## Top
            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            ## Bottom
            sBot$borderBottom <- borderStyle
            sBot$borderBottomColour <- borderColour

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol + i - 1L
              )
            ))

            ## Middle
            if (specialFormat) {
              .self$styleObjects <- append(styleObjects, list(
                list(
                  "style" = sMid,
                  "sheet" = sheet,
                  "rows" = (startRow + 1L):(startRow + nRow - 2L), # 2nd -> 2nd to last
                  "cols" = rep.int(startCol + i - 1L, nRow - 2L)
                )
              ))
            }

            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sBot,
                "sheet" = sheet,
                "rows" = startRow + nRow - 1L,
                "cols" = startCol + i - 1L
              )
            ))
          }
        } ## End of if(i == 1), i == NCol, else inside columns
      } ## End of loop through columns


      invisible(0)
    },

    rowBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType) {
      sheet <- sheet_names[[validateSheet(sheet)]]
      ## steps
      # get column class
      # get corresponding base style

      for (i in 1:nCol) {
        tmp <- genBaseColStyle(colClasses[[i]])
        sTop <- tmp$style

        ## First column
        if (i == 1) {
          if (nCol == 1) {

            ## All borders (rows and surrounding)
            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            sTop$borderBottom <- borderStyle
            sTop$borderBottomColour <- borderColour

            sTop$borderLeft <- borderStyle
            sTop$borderLeftColour <- borderColour

            sTop$borderRight <- borderStyle
            sTop$borderRightColour <- borderColour
          } else {

            ## Top, Left, Bottom
            sTop$borderTop <- borderStyle
            sTop$borderTopColour <- borderColour

            sTop$borderBottom <- borderStyle
            sTop$borderBottomColour <- borderColour

            sTop$borderLeft <- borderStyle
            sTop$borderLeftColour <- borderColour
          }
        } else if (i == nCol) {

          ## Top, Right, Bottom
          sTop$borderTop <- borderStyle
          sTop$borderTopColour <- borderColour

          sTop$borderBottom <- borderStyle
          sTop$borderBottomColour <- borderColour

          sTop$borderRight <- borderStyle
          sTop$borderRightColour <- borderColour
        } else { ## inside columns

          ## Top, Middle, Bottom
          sTop$borderTop <- borderStyle
          sTop$borderTopColour <- borderColour

          sTop$borderBottom <- borderStyle
          sTop$borderBottomColour <- borderColour
        } ## End of if(i == 1), i == NCol, else inside columns

        .self$styleObjects <- append(styleObjects, list(
          list(
            "style" = sTop,
            "sheet" = sheet,
            "rows" = (startRow):(startRow + nRow - 1L),
            "cols" = rep(startCol + i - 1L, nRow)
          )
        ))
      } ## End of loop through columns


      invisible(0)
    },

    columnBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType) {
      sheet <- sheet_names[[validateSheet(sheet)]]
      ## steps
      # get column class
      # get corresponding base style

      for (i in 1:nCol) {
        tmp <- genBaseColStyle(colClasses[[i]])
        colStyle <- tmp$style
        specialFormat <- tmp$specialFormat

        ## create style objects
        sTop <- colStyle$copy()
        sMid <- colStyle$copy()
        sBot <- colStyle$copy()

        if (nRow == 1) {

          ## Top
          sTop$borderTop <- borderStyle
          sTop$borderTopColour <- borderColour

          sTop$borderBottom <- borderStyle
          sTop$borderBottomColour <- borderColour

          sTop$borderLeft <- borderStyle
          sTop$borderLeftColour <- borderColour

          sTop$borderRight <- borderStyle
          sTop$borderRightColour <- borderColour

          .self$styleObjects <- append(styleObjects, list(
            list(
              "style" = sTop,
              "sheet" = sheet,
              "rows" = startRow,
              "cols" = startCol + i - 1L
            )
          ))
        } else {

          ## Top
          sTop$borderTop <- borderStyle
          sTop$borderTopColour <- borderColour

          sTop$borderLeft <- borderStyle
          sTop$borderLeftColour <- borderColour

          sTop$borderRight <- borderStyle
          sTop$borderRightColour <- borderColour

          ## Middle
          sMid$borderLeft <- borderStyle
          sMid$borderLeftColour <- borderColour

          sMid$borderRight <- borderStyle
          sMid$borderRightColour <- borderColour

          ## Bottom
          sBot$borderBottom <- borderStyle
          sBot$borderBottomColour <- borderColour

          sBot$borderLeft <- borderStyle
          sBot$borderLeftColour <- borderColour

          sBot$borderRight <- borderStyle
          sBot$borderRightColour <- borderColour

          colInd <- startCol + i - 1L

          .self$styleObjects <- append(styleObjects, list(
            list(
              "style" = sTop,
              "sheet" = sheet,
              "rows" = startRow,
              "cols" = colInd
            )
          ))

          if (nRow > 2) {
            .self$styleObjects <- append(styleObjects, list(
              list(
                "style" = sMid,
                "sheet" = sheet,
                "rows" = (startRow + 1L):(startRow + nRow - 2L),
                "cols" = rep(colInd, nRow - 2L)
              )
            ))
          }


          .self$styleObjects <- append(styleObjects, list(
            list(
              "style" = sBot,
              "sheet" = sheet,
              "rows" = startRow + nRow - 1L,
              "cols" = colInd
            )
          ))
        }
      } ## End of loop through columns


      invisible(0)
    },

    allBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType) {
      sheet <- sheet_names[[validateSheet(sheet)]]
      ## steps
      # get column class
      # get corresponding base style

      for (i in 1:nCol) {
        tmp <- genBaseColStyle(colClasses[[i]])
        sTop <- tmp$style

        ## All borders
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour

        sTop$borderBottom <- borderStyle
        sTop$borderBottomColour <- borderColour

        sTop$borderLeft <- borderStyle
        sTop$borderLeftColour <- borderColour

        sTop$borderRight <- borderStyle
        sTop$borderRightColour <- borderColour

        .self$styleObjects <- append(styleObjects, list(
          list(
            "style" = sTop,
            "sheet" = sheet,
            "rows" = (startRow):(startRow + nRow - 1L),
            "cols" = rep(startCol + i - 1L, nRow)
          )
        ))
      } ## End of loop through columns


      invisible(0)
    },

    setColWidths = function(sheet) {
      sheet <- validateSheet(sheet)

      widths <- colWidths[[sheet]]
      hidden <- attr(colWidths[[sheet]], "hidden", exact = TRUE)
      if (length(hidden) != length(widths)) {
        hidden <- rep("0", length(widths))
      }

      cols <- names(colWidths[[sheet]])

      autoColsInds <- widths %in% c("auto", "auto2")
      autoCols <- cols[autoColsInds]

      ## If any not auto
      if (any(!autoColsInds)) {
        widths[!autoColsInds] <- as.numeric(widths[!autoColsInds]) + 0.71
      }

      ## If any auto
      if (length(autoCols) > 0) {

        ## only run if data on worksheet
        if (worksheets[[sheet]]$sheet_data$n_elements == 0) {
          missingAuto <- autoCols
        } else if (all(is.na(worksheets[[sheet]]$sheet_data$v))) {
          missingAuto <- autoCols
        } else {

          ## First thing - get base font max character width
          baseFont <- getBaseFont()
          baseFontName <- unlist(baseFont$name, use.names = FALSE)
          if (is.null(baseFontName)) {
            baseFontName <- "calibri"
          } else {
            baseFontName <- gsub(" ", ".", tolower(baseFontName), fixed = TRUE)
            if (!baseFontName %in% names(openxlsxFontSizeLookupTable)) {
              baseFontName <- "calibri"
            }
          }

          baseFontSize <- unlist(baseFont$size, use.names = FALSE)
          if (is.null(baseFontSize)) {
            baseFontSize <- 11
          } else {
            baseFontSize <- as.numeric(baseFontSize)
            baseFontSize <- ifelse(baseFontSize < 8, 8, ifelse(baseFontSize > 36, 36, baseFontSize))
          }

          baseFontCharWidth <- openxlsxFontSizeLookupTable[[baseFontName]][baseFontSize - 7]
          allCharWidths <- rep(baseFontCharWidth, worksheets[[sheet]]$sheet_data$n_elements)
          ######### ----------------------------------------------------------------

          ## get char widths for each style object
          if (length(styleObjects) > 0 & any(!is.na(worksheets[[sheet]]$sheet_data$style_id))) {
            thisSheetName <- sheet_names[sheet]

            ## Calc font width for all styles on this worksheet
            styleIds <- worksheets[[sheet]]$sheet_data$style_id
            styObSubet <- styleObjects[sort(unique(styleIds))]
            stySubset <- lapply(styObSubet, "[[", "style")

            ## loop through stlye objects assignin a charWidth else baseFontCharWidth
            styleCharWidths <- sapply(stySubset, get_style_max_char_width, USE.NAMES = FALSE)


            ## Now assign all cells a character width
            allCharWidths <- styleCharWidths[worksheets[[sheet]]$sheet_data$style_id]
            allCharWidths[is.na(allCharWidths)] <- baseFontCharWidth
          }

          ## Now check for columns that are auto2
          auto2Inds <- which(widths %in% "auto2")
          if (length(auto2Inds) > 0 & length(worksheets[[sheet]]$mergeCells) > 0) {

            ## get cell merges
            merged_cells <- regmatches(worksheets[[sheet]]$mergeCells, regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells))

            comps <- lapply(merged_cells, function(rectCoords) unlist(strsplit(rectCoords, split = ":")))
            merge_cols <- lapply(comps, convertFromExcelRef)
            merge_cols <- lapply(merge_cols, function(x) x[x %in% cols[auto2Inds]]) ## subset to auto2Inds

            merge_rows <- lapply(comps, function(x) as.numeric(gsub("[A-Z]", "", x, perl = TRUE)))
            merge_rows <- merge_rows[sapply(merge_cols, length) > 0]
            merge_cols <- merge_cols[sapply(merge_cols, length) > 0]

            sd <- worksheets[[sheet]]$sheet_data

            if (length(merge_cols) > 0) {
              all_merged_cells <- lapply(1:length(merge_cols), function(i) {
                expand.grid(
                  "rows" = min(merge_rows[[i]]):max(merge_rows[[i]]),
                  "cols" = min(merge_cols[[i]]):max(merge_cols[[i]])
                )
              })

              all_merged_cells <- do.call("rbind", all_merged_cells)

              ## only want the sheet data in here
              refs <- paste(all_merged_cells[[1]], all_merged_cells[[2]], sep = ",")
              existing_cells <- paste(worksheets[[sheet]]$sheet_data$rows, worksheets[[sheet]]$sheet_data$cols, sep = ",")
              keep <- which(!existing_cells %in% refs & !is.na(worksheets[[sheet]]$sheet_data$v))

              sd <- SheetData$new()
              sd$cols <- worksheets[[sheet]]$sheet_data$cols[keep]
              sd$t <- worksheets[[sheet]]$sheet_data$t[keep]
              sd$v <- worksheets[[sheet]]$sheet_data$v[keep]
              sd$n_elements <- length(sd$cols)
              allCharWidths <- allCharWidths[keep]
            } else {
              sd <- worksheets[[sheet]]$sheet_data
            }
          } else {
            sd <- worksheets[[sheet]]$sheet_data
          }

          ## Now that we have the max character width for the largest font on the page calculate the column widths
          calculatedWidths <- calc_column_widths(
            sheet_data = sd,
            sharedStrings = unlist(sharedStrings, use.names = FALSE),
            autoColumns = as.integer(autoCols),
            widths = allCharWidths,
            baseFontCharWidth = baseFontCharWidth,
            minW = getOption("openxlsx.minWidth", 3),
            maxW = getOption("openxlsx.maxWidth", 250)
          )

          missingAuto <- autoCols[!autoCols %in% names(calculatedWidths)]
          widths[names(calculatedWidths)] <- calculatedWidths + 0.71
        }

        widths[missingAuto] <- 9.15
      }

      # # Check if any conflicting existing levels
      # if (any(cols %in% names(worksheets[[sheet]]$cols))) {
      #
      #   for (i in intersect(cols, names(worksheets[[sheet]]$cols))) {
      #
      #     width_hidden <- attr(colWidths[[sheet]], "hidden")[attr(colWidths[[sheet]], "names") == i]
      #     width_widths <- as.numeric(colWidths[[sheet]][attr(colWidths[[sheet]], "names") == i]) + 0.71
      #
      #     # If column already has a custom width, just update the width and hidden attributes
      #     if (grepl("customWidth", worksheets[[sheet]]$cols[[i]])) {
      #       worksheets[[sheet]]$cols[[i]] <- sub('(width=\\").*?(\\"\\shidden=\\").*?(\\")', paste0("\\1", width_widths, "\\2", width_hidden, "\\3"), worksheets[[sheet]]$cols[[i]], perl = TRUE)
      #     } else {
      #     # If column exists, but doesn't have a custom width
      #       worksheets[[sheet]]$cols[[i]] <- sub("((?<=hidden=\")(\\w)\")", paste0(width_hidden, "\" width=\"", width_widths, "\" customWidth=\"1\"/>"), worksheets[[sheet]]$cols[[i]], perl = TRUE)
      #     }
      #   }
      #
      #   cols <- cols[!cols %in% names(worksheets[[sheet]]$cols)]
      # }

      # Add remaining columns
      #if (length(cols) > 0) {
      #  colNodes <- sprintf('<col min="%s" max="%s" width="%s" hidden="%s" customWidth="1"/>', cols, cols, widths, hidden)
      #  names(colNodes) <- cols
      #  worksheets[[sheet]]$cols <- append(worksheets[[sheet]]$cols, colNodes)
      #}

    },

    writeData = function(df, sheet, startRow, startCol, colNames, colClasses, hlinkNames, keepNA, na.string, list_sep) {
      sheet <- validateSheet(sheet)
      nCols <- ncol(df)
      nRows <- nrow(df)
      df_nms <- names(df)

      allColClasses <- unlist(colClasses)
      df <- as.list(df)

      ######################################################################
      ## standardise all column types


      ## pull out NaN values
      nans <- unlist(lapply(1:nCols, function(i) {
        tmp <- df[[i]]
        if (!inherits(tmp, c("character", "list"))) {
          v <- which(is.nan(tmp) | is.infinite(tmp))
          if (length(v) == 0) {
            return(v)
          }
          return(as.integer(nCols * (v - 1) + i)) ## row position
        }
      }))

      ## convert any Dates to integers and create date style object
      if (any(c("date", "posixct", "posixt") %in% allColClasses)) {
        dInds <- which(sapply(colClasses, function(x) "date" %in% x))

        origin <- 25569L
        if (grepl('date1904="1"|date1904="true"', stri_join(unlist(workbook), collapse = ""), ignore.case = TRUE)) {
          origin <- 24107L
        }

        for (i in dInds) {
          df[[i]] <- as.integer(df[[i]]) + origin
          if (origin == 25569L){
            earlyDate <- df[[i]] < 60
            df[[i]][earlyDate] <- df[[i]][earlyDate] - 1
          }
        }

        pInds <- which(sapply(colClasses, function(x) any(c("posixct", "posixt", "posixlt") %in% x)))
        if (length(pInds) > 0 & nRows > 0) {
          parseOffset <- function(tz) {
            suppressWarnings(
              ifelse(stri_sub(tz, 1, 1) == "+", 1L, -1L)
              * (as.integer(stri_sub(tz, 2, 3)) + as.integer(stri_sub(tz, 4, 5)) / 60) / 24
            )
          }

          t <- lapply(df[pInds], function(x) format(x, "%z"))
          offSet <- lapply(t, parseOffset)
          offSet <- lapply(offSet, function(x) ifelse(is.na(x), 0, x))

          for (i in 1:length(pInds)) {
            df[[pInds[i]]] <- as.numeric(as.POSIXct(df[[pInds[i]]])) / 86400 + origin + offSet[[i]]
          }
        }
      }


      ## convert any Dates to integers and create date style object
      if (any(c("currency", "accounting", "percentage", "3", "comma") %in% allColClasses)) {
        cInds <- which(sapply(colClasses, function(x) any(c("accounting", "currency", "percentage", "3", "comma") %in% tolower(x))))
        for (i in cInds) {
          df[[i]] <- as.numeric(gsub("[^0-9\\.-]", "", df[[i]], perl = TRUE))
        }
        class(df[[i]]) <- "numeric"
      }

      ## convert scientific
      if ("scientific" %in% allColClasses) {
        for (i in which(sapply(colClasses, function(x) "scientific" %in% x))) {
          class(df[[i]]) <- "numeric"
        }
      }

      ##
      if ("list" %in% allColClasses) {
        for (i in which(sapply(colClasses, function(x) "list" %in% x))) {
          df[[i]] <- sapply(lapply(df[[i]], unlist), stri_join, collapse = list_sep)
        }
      }

      if (any(c("formula", "array_formula") %in% allColClasses)) {

        frm <- "formula"
        cls <- "openxlsx_formula"

        if ("array_formula" %in% allColClasses) {
          frm <- "array_formula"
          cls <- "openxlsx_array_formula"
        }

        for (i in which(sapply(colClasses, function(x) frm %in% x))) {
          df[[i]] <- replaceIllegalCharacters(as.character(df[[i]]))
          class(df[[i]]) <- cls
        }
      }

      if ("hyperlink" %in% allColClasses) {
        for (i in which(sapply(colClasses, function(x) "hyperlink" %in% x))) {
          class(df[[i]]) <- "hyperlink"
        }
      }

      colClasses <- sapply(df, function(x) tolower(class(x))[[1]]) ## by here all cols must have a single class only


      ## convert logicals (Excel stores logicals as 0 & 1)
      if ("logical" %in% allColClasses) {
        for (i in which(sapply(colClasses, function(x) "logical" %in% x))) {
          class(df[[i]]) <- "numeric"
        }
      }

      ## convert all numerics to character (this way preserves digits)
      if ("numeric" %in% colClasses) {
        for (i in which(sapply(colClasses, function(x) "numeric" %in% x))) {
          class(df[[i]]) <- "character"
        }
      }


      ## End standardise all column types
      ######################################################################


      ## cell types
      t <- build_cell_types_integer(classes = colClasses, n_rows = nRows)

      for (i in which(sapply(colClasses, function(x) !"character" %in% x & !"numeric" %in% x))) {
        df[[i]] <- as.character(df[[i]])
      }

      ## cell values
      v <- as.character(t(as.matrix(
        data.frame(df, stringsAsFactors = FALSE, check.names = FALSE, fix.empty.names = FALSE)
      )))


      if (keepNA) {
        if (is.null(na.string)) {
          # t[is.na(v)] <- 4L
          v[is.na(v)] <- "#N/A"
        } else {
          # t[is.na(v)] <- 1L
          v[is.na(v)] <- as.character(na.string)
        }
      } else {
        t[is.na(v)] <- as.character(NA)
        v[is.na(v)] <- as.character(NA)
      }

      ## If any NaN values
      if (length(nans) > 0) {
        t[nans] <- 4L
        v[nans] <- "#NUM!"
      }


      # prepend column headers
      if (colNames) {
        t <- c(rep.int(1L, nCols), t)
        v <- c(df_nms, v)
        nRows <- nRows + 1L
      }


      ## Formulas
      f_in <- rep.int(as.character(NA), length(t))
      any_functions <- FALSE
      ref_cell <- paste0(int_2_cell_ref(startCol), startRow)

      if (any(c("openxlsx_formula", "openxlsx_array_formula") %in% colClasses)) {

        ## alter the elements of t where we have a formula to be "str"
        if ("openxlsx_formula" %in% colClasses) {
          formula_cols <- which(sapply(colClasses, function(x) "openxlsx_formula" %in% x, USE.NAMES = FALSE), useNames = FALSE)
          formula_strs <- stri_join("<f>", unlist(df[formula_cols], use.names = FALSE), "</f>")
        } else { # openxlsx_array_formula
          formula_cols <- which(sapply(colClasses, function(x) "openxlsx_array_formula" %in% x, USE.NAMES = FALSE), useNames = FALSE)
          formula_strs <- stri_join("<f t=\"array\" ref=\"", ref_cell, ":", ref_cell, "\">", unlist(df[formula_cols], use.names = FALSE), "</f>")
        }
        formula_inds <- unlist(lapply(formula_cols, function(i) i + (1:(nRows - colNames) - 1) * nCols + (colNames * nCols)), use.names = FALSE)
        f_in[formula_inds] <- formula_strs
        any_functions <- TRUE

        rm(formula_cols)
        rm(formula_strs)
        rm(formula_inds)
      }

      suppressWarnings(try(rm(df), silent = TRUE))

      ## Append hyperlinks, convert h to s in cell type
      hyperlink_cols <- which(sapply(colClasses, function(x) "hyperlink" %in% x, USE.NAMES = FALSE), useNames = FALSE)
      if (length(hyperlink_cols) > 0) {
        hyperlink_inds <- sort(unlist(lapply(hyperlink_cols, function(i) i + (1:(nRows - colNames) - 1) * nCols + (colNames * nCols)), use.names = FALSE))
        na_hyperlink <- intersect(hyperlink_inds, which(is.na(t)))

        if (length(hyperlink_inds) > 0) {
          t[t %in% 9] <- 1L ## set cell type to "s"

          hyperlink_refs <- convert_to_excel_ref_expand(cols = hyperlink_cols + startCol - 1, LETTERS = LETTERS, rows = as.character((startRow + colNames):(startRow + nRows - 1L)))

          if (length(na_hyperlink) > 0) {
            to_remove <- which(hyperlink_inds %in% na_hyperlink)
            hyperlink_refs <- hyperlink_refs[-to_remove]
            hyperlink_inds <- hyperlink_inds[-to_remove]
          }

          exHlinks <- worksheets[[sheet]]$hyperlinks
          targets <- replaceIllegalCharacters(v[hyperlink_inds])

          if (!is.null(hlinkNames) & length(hlinkNames) == length(hyperlink_inds)) {
            v[hyperlink_inds] <- hlinkNames
          } ## this is text to display instead of hyperlink

          ## create hyperlink objects
          newhl <- lapply(1:length(hyperlink_inds), function(i) {
            Hyperlink$new(ref = hyperlink_refs[i], target = targets[i], location = NULL, display = NULL, is_external = TRUE)
          })

          .self$worksheets[[sheet]]$hyperlinks <- append(worksheets[[sheet]]$hyperlinks, newhl)
        }
      }


      ## convert all strings to references in sharedStrings and update values (v)
      strFlag <- which(t == 1L)
      newStrs <- v[strFlag]
      if (length(newStrs) > 0) {
        newStrs <- replaceIllegalCharacters(newStrs)
        newStrs <- stri_join("<si><t xml:space=\"preserve\">", newStrs, "</t></si>")

        uNewStr <- unique(newStrs)

        .self$updateSharedStrings(uNewStr)
        v[strFlag] <- match(newStrs, sharedStrings) - 1L
      }

      # ## Create cell list of lists
      worksheets[[sheet]]$sheet_data$write(
        rows_in = startRow:(startRow + nRows - 1L),
        cols_in = startCol:(startCol + nCols - 1L),
        t_in = t,
        v_in = v,
        f_in = f_in,
        any_functions = any_functions
      )



      invisible(0)
    }
  )

)
