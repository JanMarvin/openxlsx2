
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
    "sheetOrder" = "integer",

    # allows path to be set during initiation or later
    "path" = "character",
    "styleObjectsList" = "list"
  ),

  methods = list(

    initialize = function(
      creator = "",
      title = NULL,
      subject = NULL,
      category = NULL
    ) {
      .self$charts <- list()
      .self$isChartSheet <- logical()

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

      .self$sheet_names <- character()
      .self$sheetOrder <- integer()

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

      if (length(path)) {
        .self$path <- path
      }

      # FIXME styleObjectsList() may be getting removed [11]
      .self$styleObjectsList <- list()

      invisible(.self)
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
      newSheetIndex <- length(.self$worksheets) + 1L

      if (newSheetIndex > 1) {
        sheetId <-
          max(as.integer(regmatches(
            workbook$sheets,
            regexpr('(?<=sheetId=")[0-9]+', workbook$sheets, perl = TRUE)
          ))) + 1L
      } else {
        sheetId <- 1
      }


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
          .self$worksheets,
          Worksheet$new(
            showGridLines = showGridLines,
            tabSelected   = newSheetIndex == 1,
            tabColour     = tabColour,
            zoom          = zoom,
            oddHeader     = oddHeader,
            oddFooter     = oddFooter,
            evenHeader    = evenHeader,
            evenFooter    = evenFooter,
            firstHeader   = firstHeader,
            firstFooter   = firstFooter,
            paperSize     = paperSize,
            orientation   = orientation,
            hdpi          = hdpi,
            vdpi          = vdpi
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

      .self$sheetOrder <- c(.self$sheetOrder, as.integer(newSheetIndex))
      .self$sheet_names <- c(.self$sheet_names, sheetName)

      # Jordan is a little worried this may change something
      # invisible(newSheetIndex)
      invisible(.self)
    },

    cloneWorksheet = function(sheetName, clonedSheet) {
      clonedSheet <- .self$validateSheet(clonedSheet)
      if (!missing(sheetName)) {
        if (grepl(pattern = ":", x = sheetName)) {
          stop("colon not allowed in sheet names in Excel")
        }
      }
      newSheetIndex <- length(.self$worksheets) + 1L
      if (newSheetIndex > 1) {
        sheetId <-
          max(as.integer(regmatches(
            .self$workbook$sheets,
            regexpr('(?<=sheetId=")[0-9]+', .self$workbook$sheets, perl = TRUE)
          ))) + 1L
      } else {
        sheetId <- 1
      }


      ## copy visibility from cloned sheet!
      visible <-
        regmatches(
          .self$workbook$sheets[[clonedSheet]],
          regexpr('(?<=state=")[^"]+', .self$workbook$sheets[[clonedSheet]], perl = TRUE)
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
        append(.self$worksheets, .self$worksheets[[clonedSheet]]$copy())


      ## update content_tyes
      ## add a drawing.xml for the worksheet
      .self$Content_Types <-
        c(
          .self$Content_Types,
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
        .self$workbook.xml.rels,
        sprintf(
          '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>',
          newSheetIndex
        )
      )

      ## create sheet.rels to simplify id assignment
      .self$worksheets_rels[[newSheetIndex]] <-
        genBaseSheetRels(newSheetIndex)
      .self$drawings_rels[[newSheetIndex]] <- .self$drawings_rels[[clonedSheet]]

      # give each chart its own filename (images can re-use the same file, but charts can't)
      .self$drawings_rels[[newSheetIndex]] <-
        sapply(.self$drawings_rels[[newSheetIndex]], function(rl) {
          chartfiles <-
            regmatches(
              rl,
              gregexpr("(?<=charts/)chart[0-9]+\\.xml", rl, perl = TRUE)
            )[[1]]
          for (cf in chartfiles) {
            chartid <- length(.self$charts) + 1
            newname <- stri_join("chart", chartid, ".xml")
            fl <- .self$charts[cf]

            # Read the chartfile and adjust all formulas to point to the new
            # sheet name instead of the clone source
            # The result is saved to a new chart xml file
            newfl <- file.path(dirname(fl), newname)
            .self$charts[newname] <- newfl
            chart <- read_xml(fl, pointer = FALSE)
            chart <-
              gsub(
                stri_join("(?<=')", .self$sheet_names[[clonedSheet]], "(?='!)"),
                stri_join("'", sheetName, "'"),
                chart,
                perl = TRUE
              )
            chart <-
              gsub(
                stri_join("(?<=[^A-Za-z0-9])", .self$sheet_names[[clonedSheet]], "(?=!)"),
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
      .self$drawings[[newSheetIndex]] <- .self$drawings[[clonedSheet]]

      .self$vml_rels[[newSheetIndex]] <- .self$vml_rels[[clonedSheet]]
      .self$vml[[newSheetIndex]] <- .self$vml[[clonedSheet]]

      .self$isChartSheet[[newSheetIndex]] <- .self$isChartSheet[[clonedSheet]]
      .self$comments[[newSheetIndex]] <- .self$comments[[clonedSheet]]
      .self$threadComments[[newSheetIndex]] <- .self$threadComments[[clonedSheet]]

      .self$rowHeights[[newSheetIndex]] <- .self$rowHeights[[clonedSheet]]
      .self$colWidths[[newSheetIndex]] <- .self$colWidths[[clonedSheet]]

      .self$colOutlineLevels[[newSheetIndex]] <- .self$colOutlineLevels[[clonedSheet]]
      .self$outlineLevels[[newSheetIndex]] <- .self$outlineLevels[[clonedSheet]]

      .self$sheetOrder <- c(.self$sheetOrder, as.integer(newSheetIndex))
      .self$sheet_names <- c(.self$sheet_names, sheetName)


      ############################
      ## STYLE
      ## ... objects are stored in a global list, so we need to get all styles
      ## assigned to the cloned sheet and duplicate them

      # TODO can we replace Filter() and Map()?
      sheetStyles <- Filter(function(s) {
        s$sheet == sheet_names[[clonedSheet]]
      }, .self$styleObjects)
      .self$styleObjects <- c(
        .self$styleObjects,
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

      tbls <- .self$tables[attr(.self$tables, "sheet") == clonedSheet]
      for (t in tbls) {
        # Extract table name, displayName and ID from the xml
        oldname <- regmatches(t, regexpr('(?<= name=")[^"]+', t, perl = TRUE))
        olddispname <- regmatches(t, regexpr('(?<= displayName=")[^"]+', t, perl = TRUE))
        oldid <- regmatches(t, regexpr('(?<= id=")[^"]+', t, perl = TRUE))
        ref <- regmatches(t, regexpr('(?<= ref=")[^"]+', t, perl = TRUE))

        # Find new, unused table names by appending _n, where n=1,2,...
        n <- 0
        while (stri_join(oldname, "_", n) %in% attr(.self$tables, "tableName")) {
          n <- n + 1
        }
        newname <- stri_join(oldname, "_", n)
        newdispname <- stri_join(olddispname, "_", n)
        newid <- as.character(length(.self$tables) + 3L)

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

        oldtables <- .self$tables
        .self$tables <- c(oldtables, newt)
        names(.self$tables) <- c(names(oldtables), ref)
        attr(.self$tables, "sheet") <-
          c(attr(oldtables, "sheet"), newSheetIndex)
        attr(.self$tables, "tableName") <-
          c(attr(oldtables, "tableName"), newname)

        oldparts <- .self$worksheets[[newSheetIndex]]$tableParts
        .self$worksheets[[newSheetIndex]]$tableParts <-
          c(oldparts, sprintf('<tablePart r:id="rId%s"/>', newid))
        attr(.self$worksheets[[newSheetIndex]]$tableParts, "tableName") <-
          c(attr(oldparts, "tableName"), newname)
        names(attr(.self$worksheets[[newSheetIndex]]$tableParts, "tableName")) <-
          c(names(attr(oldparts, "tableName")), ref)

        .self$Content_Types <-
          c(
            .self$Content_Types,
            sprintf(
              '<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
              newid
            )
          )
        .self$tables.xml.rels <- append(.self$tables.xml.rels, "")

        .self$worksheets_rels[[newSheetIndex]] <-
          c(
            .self$worksheets_rels[[newSheetIndex]],
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

      # invisible(newSheetIndex)
      invisible(.self)
    },

    addChartSheet = function(sheetName,
      tabColour = NULL,
      zoom = 100) {
      # TODO private$new_sheet_index()?
      newSheetIndex <- length(.self$worksheets) + 1L

      if (newSheetIndex > 1) {
        sheetId <-
          max(as.integer(regmatches(
            .self$workbook$sheets,
            regexpr('(?<=sheetId=")[0-9]+', .self$workbook$sheets, perl = TRUE)
          ))) + 1L
      } else {
        sheetId <- 1
      }

      ##  Add sheet to workbook.xml
      .self$workbook$sheets <-
        c(
          .self$workbook$sheets,
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
          .self$worksheets,
          ChartSheet$new(
            tabSelected = newSheetIndex == 1,
            tabColour = tabColour,
            zoom = zoom
          )
        )
      .self$sheet_names <- c(.self$sheet_names, sheetName)

      ## update content_tyes
      .self$Content_Types <-
        c(
          .self$Content_Types,
          sprintf(
            '<Override PartName="/xl/chartsheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>',
            newSheetIndex
          )
        )

      ## Update xl/rels
      .self$workbook.xml.rels <- c(
        .self$workbook.xml.rels,
        sprintf(
          '<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet%s.xml"/>',
          newSheetIndex
        )
      )



      ## add a drawing.xml for the worksheet
      .self$Content_Types <-
        c(
          .self$Content_Types,
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

      .self$sheetOrder <- c(.self$sheetOrder, newSheetIndex)

      # invisible(newSheetIndex)
      invisible(.self)
    },

    # TODO saveWorkbook can be shortened a lot by some formatting and by using a
    # function that creates all the temporary directories and subdirectries as a
    # named list
    saveWorkbook = function(path = .self$path, overwrite = TRUE) {
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

      .self$preSaveCleanUp()

      nSheets <- length(.self$worksheets)
      nThemes <- length(.self$theme)
      nPivots <- length(.self$pivotDefinitions)
      nSlicers <- length(.self$slicers)
      nComments <- sum(lengths(.self$comments) > 0)
      nThreadComments <- sum(lengths(.self$threadComments) > 0)
      nPersons <- length(.self$persons)
      nVML <- sum(lengths(.self$vml) > 0)

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

      if (length(.self$media)) {
        xlmediaDir <- file.path(tmpDir, "xl", "media")
        dir.create(path = xlmediaDir, recursive = TRUE)
      }


      ## will always have a theme
      xlthemeDir <- file.path(tmpDir, "xl", "theme")
      dir.create(path = xlthemeDir, recursive = TRUE)

      if (is.null(.self$theme)) {
        con <- file(file.path(xlthemeDir, "theme1.xml"), open = "wb")
        writeBin(charToRaw(genBaseTheme()), con)
        close(con)
      } else {
        # TODO replace with seq_len() or seq_along()
        lapply(seq_len(nThemes), function(i) {
          con <-
            file(file.path(xlthemeDir, stri_join("theme", i, ".xml")), open = "wb")
          writeBin(charToRaw(pxml(.self$theme[[i]])), con)
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
      if (length(.self$charts)) {
        file.copy(
          from = dirname(.self$charts[1]),
          to = file.path(tmpDir, "xl"),
          recursive = TRUE
        )
      }


      ## xl/comments.xml
      if (nComments > 0 | nVML > 0) {
        # TODO use seq_len() or seq_along()?
        for (i in seq_len(nSheets)) {
          if (length(comments[[i]])) {
            fn <- sprintf("comments%s.xml", i)

            .self$Content_Types <- c(
              .self$Content_Types,
              sprintf(
                '<Override PartName="/xl/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>',
                fn
              )
            )

            .self$worksheets_rels[[i]] <- unique(c(
              .self$worksheets_rels[[i]],
              sprintf(
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../%s"/>',
                fn
              )
            ))

            writeCommentXML(
              comment_list = .self$comments[[i]],
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
          if (length(.self$threadComments[[i]])) {
            fl <- .self$threadComments[[i]]
            file.copy(
              from = fl,
              to = file.path(xlThreadComments, basename(fl)),
              overwrite = TRUE,
              copy.date = TRUE
            )

            .self$worksheets_rels[[i]] <- unique(c(
              .self$worksheets_rels[[i]],
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
          from = .self$persons,
          to = personDir,
          overwrite = TRUE
        )

      }



      if (length(.self$embeddings)) {
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
        # TODO consider just making a function to create a bunch of directories
        # and return as a named list?  Easier/cleaner than checking for each
        # element if we just go seq_along()?
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

        for (i in seq_along(.self$pivotTables)) {
          file.copy(
            from = .self$pivotTables[i],
            to = file.path(pivotTablesDir, sprintf("pivotTable%s.xml", i)),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }

        for (i in seq_along(.self$pivotDefinitions)) {
          file.copy(
            from = .self$pivotDefinitions[i],
            to = file.path(pivotCacheDir, sprintf("pivotCacheDefinition%s.xml", i)),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }

        for (i in seq_along(.self$pivotRecords)) {
          file.copy(
            from = .self$pivotRecords[i],
            to = file.path(pivotCacheDir, sprintf("pivotCacheRecords%s.xml", i)),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }

        for (i in seq_along(.self$pivotDefinitionsRels)) {
          file.copy(
            from = .self$pivotDefinitionsRels[i],
            to = file.path(
              pivotCacheRelsDir,
              sprintf("pivotCacheDefinition%s.xml.rels", i)
            ),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }

        for (i in seq_along(.self$pivotTables.xml.rels)) {
          write_file(
            body = .self$pivotTables.xml.rels[[i]],
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

        for (i in seq_along(.self$slicers)) {
          # TODO consider nzchar()?
          if (nchar(.self$slicers[i])) {
            file.copy(from = .self$slicers[i], to = file.path(slicersDir, sprintf("slicer%s.xml", i)))
          }
        }



        for (i in seq_along(.self$slicerCaches)) {
          write_file(
            body = .self$slicerCaches[[i]],
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
        body = pxml(.self$core),
        tail = "",
        fl = file.path(docPropsDir, "core.xml")
      )

      ## write workbook.xml.rels
      write_file(
        head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        body = pxml(.self$workbook.xml.rels),
        tail = "</Relationships>",
        fl = file.path(xlrelsDir, "workbook.xml.rels")
      )

      ## write tables

      # TODO remove length() check since we have seq_along()
      if (length(unlist(.self$tables, use.names = FALSE))) {
        for (i in seq_along(unlist(.self$tables, use.names = FALSE))) {
          if (!grepl("openxlsx_deleted", attr(.self$tables, "tableName")[i], fixed = TRUE)) {
            write_file(
              body = pxml(unlist(.self$tables, use.names = FALSE)[[i]]),
              fl = file.path(xlTablesDir, sprintf("table%s.xml", i + 2))
            )
            if (.self$tables.xml.rels[[i]] != "") {
              write_file(
                body = .self$tables.xml.rels[[i]],
                fl = file.path(xlTablesRelsDir, sprintf("table%s.xml.rels", i + 2))
              )
            }
          }
        }
      }


      ## write query tables
      if (length(.self$queryTables)) {
        xlqueryTablesDir <- file.path(tmpDir, "xl", "queryTables")
        dir.create(path = xlqueryTablesDir, recursive = TRUE)

        for (i in seq_along(.self$queryTables)) {
          write_file(
            body = .self$queryTables[[i]],
            fl = file.path(xlqueryTablesDir, sprintf("queryTable%s.xml", i))
          )
        }
      }

      ## connections
      if (length(.self$connections)) {
        write_file(body = .self$connections, fl = file.path(xlDir, "connections.xml"))
      }

      ## externalLinks
      if (length(.self$externalLinks)) {
        externalLinksDir <- file.path(tmpDir, "xl", "externalLinks")
        dir.create(path = externalLinksDir, recursive = TRUE)

        for (i in seq_along(.self$externalLinks)) {
          write_file(
            body = .self$externalLinks[[i]],
            fl = file.path(externalLinksDir, sprintf("externalLink%s.xml", i))
          )
        }
      }

      ## externalLinks rels
      if (length(.self$externalLinksRels)) {
        externalLinksRelsDir <-
          file.path(tmpDir, "xl", "externalLinks", "_rels")
        dir.create(path = externalLinksRelsDir, recursive = TRUE)

        for (i in seq_along(.self$externalLinksRels)) {
          write_file(
            body = .self$externalLinksRels[[i]],
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
      # for (i in seq_len(nSheets)) {
      #   writeLines(genPrinterSettings(), file.path(printDir, sprintf("printerSettings%s.bin", i)))
      # }

      ## media (copy file from origin to destination)
      # TODO replace with seq_along()
      for (x in .self$media) {
        file.copy(x, file.path(xlmediaDir, names(.self$media)[which(.self$media == x)]))
      }

      ## VBA Macro
      if (!is.null(.self$vbaProject)) {
        file.copy(.self$vbaProject, xlDir)
      }

      ## write worksheet, worksheet_rels, drawings, drawing_rels
      .self$writeSheetDataXML(
        xldrawingsDir,
        xldrawingsRelsDir,
        xlworksheetsDir,
        xlworksheetsRelsDir
      )

      ## write sharedStrings.xml
      ct <- .self$Content_Types
      if (length(.self$sharedStrings)) {
        write_file(
          head = sprintf(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%s" uniqueCount="%s">',
            length(.self$sharedStrings),
            attr(.self$sharedStrings, "uniqueCount")
          ),
          #body = stri_join(set_sst(attr(sharedStrings, "text")), collapse = "", sep = " "),
          body = stri_join(.self$sharedStrings, collapse = "", sep = " "),
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


      styleXML <- .self$styles
      styleXML$numFmts <-
        stri_join(
          sprintf('<numFmts count="%s">', length(.self$styles$numFmts)),
          pxml(.self$styles$numFmts),
          "</numFmts>"
        )
      styleXML$fonts <-
        stri_join(
          sprintf('<fonts count="%s">', length(.self$styles$fonts)),
          pxml(.self$styles$fonts),
          "</fonts>"
        )
      styleXML$fills <-
        stri_join(
          sprintf('<fills count="%s">', length(.self$styles$fills)),
          pxml(.self$styles$fills),
          "</fills>"
        )
      styleXML$borders <-
        stri_join(
          sprintf('<borders count="%s">', length(.self$styles$borders)),
          pxml(.self$styles$borders),
          "</borders>"
        )
      styleXML$cellStyleXfs <-
        c(
          sprintf('<cellStyleXfs count="%s">', length(.self$styles$cellStyleXfs)),
          pxml(.self$styles$cellStyleXfs),
          "</cellStyleXfs>"
        )
      styleXML$cellXfs <-
        stri_join(
          sprintf('<cellXfs count="%s">', length(.self$styles$cellXfs)),
          paste0(.self$styles$cellXfs, collapse = ""),
          "</cellXfs>"
        )
      styleXML$cellStyles <-
        stri_join(
          sprintf('<cellStyles count="%s">', length(.self$styles$cellStyles)),
          pxml(.self$styles$cellStyles),
          "</cellStyles>"
        )
      # TODO
      # tableStyles
      # extLst

      # TODO replace ifelse() with just if () else
      styleXML$dxfs <-
        ifelse(
          length(.self$styles$dxfs) == 0,
          '<dxfs count="0"/>',
          stri_join(
            sprintf('<dxfs count="%s">', length(.self$styles$dxfs)),
            stri_join(unlist(.self$styles$dxfs), sep = " ", collapse = ""),
            "</dxfs>"
          )
        )

      ## write styles.xml
      #if(class(.self$styles_xml) == "uninitializedField") {
        write_file(
          head = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">',
          body = pxml(styleXML),
          tail = "</styleSheet>",
          fl = file.path(xlDir, "styles.xml")
        )
      #} else {
      #  write_file(
      #    head = '',
      #    body = .self$styles_xml,
      #    tail = '',
      #    fl = file.path(xlDir, "styles.xml")
      #  )
      #}

      ## write workbook.xml
      workbookXML <- .self$workbook
      workbookXML$sheets <-
        stri_join("<sheets>", pxml(workbookXML$sheets), "</sheets>")
      if (length(workbookXML$definedNames)) {
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
        .self$workbook$sheets[order(.self$sheetOrder)] ## Need to reset sheet order to allow multiple savings

      ## compress to xlsx

      tmpFile <- tempfile(
        tmpdir = tmpDir,
        # TODO make .self$vbaProject be TRUE/FALSE
        fileext = if (isTRUE(.self$vbaProject)) ".xlsm" else ".xlsx"
      )

      ## zip it
      zip::zip(
        zipfile = tmpFile,
        files = list.files(tmpDir, full.names = FALSE),
        recurse = TRUE,
        compression_level = getOption("openxlsx.compresssionevel", 6),
        include_directories = FALSE,
        # change the working directory for this
        root = tmpDir,
        # change default to match historical zipr
        mode = "cherry-pick"
      )

      # reset styles - maintain any changes to base font
      # TODO: why would I want to do that?
      baseFont <- .self$styles$fonts[[1]]
      .self$styles <-
        genBaseStyleSheet(.self$styles$dxfs,
          tableStyles = .self$styles$tableStyles,
        )
      .self$styles$fonts[[1]] <- baseFont


      # Copy file; stop if filed
      if (!file.copy(from = tmpFile, to = path, overwrite = overwrite)) {
        stop("Failed to save workbook")
      }

      # (re)assign file path (if successful)
      .self$path <- path
      invisible(.self)
    },

    updateSharedStrings = function(uNewStr) {
      ## Function will return named list of references to new strings
      uStr <- uNewStr[which(!uNewStr %in% .self$sharedStrings)]
      uCount <- attr(.self$sharedStrings, "uniqueCount")
      .self$sharedStrings <- append(.self$sharedStrings, uStr)

      attr(.self$sharedStrings, "uniqueCount") <- uCount + length(uStr)
      invisible(.self)
    },

    # Doesn't make any assignments, could be pulled out
    validateSheet = function(sheetName) {
      if (!is.numeric(sheetName)) {
        if (is.null(.self$sheet_names)) {
          stop("Workbook does not contain any worksheets.", call. = FALSE)
        }
      }

      if (is.numeric(sheetName)) {
        if (sheetName > length(.self$sheet_names)) {
          stop(sprintf("This Workbook only has %s sheets.", length(.self$sheet_names)),
            call. =
              FALSE
          )
        }

        # TODO consider return(invisible(.self))
        return(sheetName)
      } else if (!sheetName %in% replaceXMLEntities(.self$sheet_names)) {
        stop(sprintf("Sheet '%s' does not exist.", replaceXMLEntities(sheetName)), call. = FALSE)
      }

      return(which(replaceXMLEntities(sheet_names) == sheetName))
    },

    # TODO Does this need to be checked?  No sheet name can be NA right?
    # res <- .self$sheet_names[ind]; stopifnot(!anyNA(ind))

    # Doesn't make any assignments, could be pulled out
    getSheetName = function(sheetIndex) {
      if (any(length(.self$sheet_names) < sheetIndex)) {
        stop(sprintf("Workbook only contains %s sheet(s).", length(.self$sheet_names)))
      }

      .self$sheet_names[sheetIndex]
    },

    buildTable = function(sheet,
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
      showColumnStripes = 0) {
      ## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
      id <- as.character(length(.self$tables) + 3L)
      sheet <- .self$validateSheet(sheet)

      ## build table XML and save to tables field
      table <-
        sprintf(
          '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="%s" name="%s" displayName="%s" ref="%s" totalsRowCount="%s"',
          id,
          tableName,
          tableName,
          ref,
          # TODO force as.integer() for all values
          as.integer(totalsRowCount)
        )
      # because tableName might be native encoded non-ASCII strings, we need to ensure
      # it's UTF-8 encoded
      table <- enc2utf8(table)

      nms <- names(.self$tables)
      tSheets <- attr(.self$tables, "sheet")
      tNames <- attr(.self$tables, "tableName")

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
          .self$tables,
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
          .self$worksheets[[sheet]]$tableParts,
          sprintf('<tablePart r:id="rId%s"/>', id)
        )
      attr(.self$worksheets[[sheet]]$tableParts, "tableName") <-
        c(tNames[tSheets == sheet &
            !grepl("openxlsx_deleted", tNames, fixed = TRUE)], tableName)



      ## update Content_Types
      .self$Content_Types <-
        c(          .self$Content_Types,
          sprintf(
            '<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
            id
          )
        )

      ## create a table.xml.rels
      .self$tables.xml.rels <- append(.self$tables.xml.rels, "")

      ## update worksheets_rels
      .self$worksheets_rels[[sheet]] <- c(
        .self$worksheets_rels[[sheet]],
        sprintf(
          '<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',
          id,
          id
        )
      )

      invisible(.self)
    },

    writeDrawingVML = function(dir) {
      for (i in seq_along(.self$comments)) {
        id <- 1025

        cd <- unlist(lapply(.self$comments[[i]], "[[", "clientData"))
        nComments <- length(cd)

        ## write head
        if (nComments > 0 | length(.self$vml[[i]])) {
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

        if (length(.self$vml[[i]])) {
          write(
            x = .self$vml[[i]],
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            append = TRUE
          )
        }

        # TODO nComments and .self$vml is already checked
        if (nComments > 0 | length(.self$vml[[i]])) {
          write(
            x = "</xml>",
            file = file.path(dir, sprintf("vmlDrawing%s.vml", i)),
            append = TRUE
          )
          .self$worksheets[[i]]$legacyDrawing <-
            '<legacyDrawing r:id="rId2"/>'
        }

      }

      for (i in seq_along(.self$drawings_vml)) {
        write(
          x = .self$drawings_vml[[i]],
          file = file.path(dir, sprintf("vmlDrawing%s.vml", i))
        )
      }

      invisible(.self)
    },

    updateStyles = function(style) {
      # TODO assert_class(style, "Style")

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
        if (as.integer(style$numFmt$numFmtId)) {
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
        fillNode <- createFillNode(style)
        if (!is.null(fillNode)) {
          fillId <- which(.self$styles$fills == fillNode) - 1L

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
        borderNode <- createBorderNode(style)
        borderId <- which(.self$styles$borders == borderNode) - 1L

        if (length(borderId) == 0) {
          borderId <- length(.self$styles$borders)
          .self$styles$borders <- c(.self$styles$borders, borderNode)
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

      if (length(childNodes)) {
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

      styleId <- which(.self$styles$cellXfs == xfNode) - 1L
      if (length(styleId) == 0) {
        styleId <- length(.self$styles$cellXfs)
        .self$styles$cellXfs <- c(.self$styles$cellXfs, xfNode)
      }


      # Seems to be fine to return .self
      # return(as.integer(styleId))
      invisible(.self)
    },

    updateCellStyles = function() {
      flag <- TRUE
      for (style in .self$cellStyleObjects) {
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
          fontId <- which(.self$styles$font == fontNode) - 1L

          if (length(fontId) == 0) {
            fontId <- length(.self$styles$fonts)
            .self$styles$fonts <- append(.self$styles[["fonts"]], fontNode)
          }

          xfNode$fontId <- fontId
          xfNode <- append(xfNode, list("applyFont" = "1"))
        }


        ## numFmt
        if (!is.null(style$numFmt)) {
          if (as.integer(style$numFmt$numFmtId)) {
            numFmtId <- style$numFmt$numFmtId
            if (as.integer(numFmtId) > 163L) {
              tmp <- style$numFmt$formatCode

              .self$styles$numFmts <- unique(c(
                .self$styles$numFmts,
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
          fillNode <- createFillNode(style)
          if (!is.null(fillNode)) {
            fillId <- which(.self$styles$fills == fillNode) - 1L

            if (length(fillId) == 0) {
              fillId <- length(.self$styles$fills)
              .self$styles$fills <- c(.self$styles$fills, fillNode)
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
          borderNode <- createBorderNode(style)
          borderId <- which(.self$styles$borders == borderNode) - 1L

          if (length(borderId) == 0) {
            borderId <- length(.self$styles$borders)
            .self$styles$borders <- c(.self$styles$borders, borderNode)
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
          .self$styles$cellStyleXfs <- c(.self$styles$cellStyleXfs, xfNode)
        }
      }

      invisible(.self)
    },

    # Nothing is assigned to .self, so this is fine to not return .self
    createFontNode = function(style) {
      # TODO assert_class(style, "Style")
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

    # TODO can this just be saved as an private element?
    getBaseFont = function() {
      baseFont <- .self$styles$fonts[[1]]

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

    setSheetName = function(sheet, newSheetName) {
      # TODO assert sheet class?
      if (newSheetName %in% .self$sheet_names) {
        stop(sprintf("Sheet %s already exists!", newSheetName))
      }

      sheet <- .self$validateSheet(sheet)

      oldName <- .self$sheet_names[[sheet]]
      .self$sheet_names[[sheet]] <- newSheetName

      ## Rename in workbook
      sheetId <-
        regmatches(
          .self$workbook$sheets[[sheet]],
          regexpr('(?<=sheetId=")[0-9]+', .self$workbook$sheets[[sheet]], perl = TRUE)
        )
      rId <-
        regmatches(
          workbook$sheets[[sheet]],
          regexpr('(?<= r:id="rId)[0-9]+', .self$workbook$sheets[[sheet]], perl = TRUE)
        )
      .self$workbook$sheets[[sheet]] <-
        sprintf(
          '<sheet name="%s" sheetId="%s" r:id="rId%s"/>',
          newSheetName,
          sheetId,
          rId
        )

      ## rename styleObjects sheet component
      if (length(styleObjects)) {
        .self$styleObjects <- lapply(styleObjects, function(x) {
          if (x$sheet == oldName) {
            x$sheet <- newSheetName
          }

          return(x)
        })
      }

      ## rename defined names
      if (length(.self$workbook$definedNames)) {
        belongTo <- getDefinedNamesSheet(.self$workbook$definedNames)
        toChange <- belongTo == oldName
        if (any(toChange)) {
          newSheetName <- sprintf("'%s'", newSheetName)
          tmp <-
            gsub(oldName, newSheetName, .self$workbook$definedName[toChange], fixed = TRUE)
          tmp <- gsub("'+", "'", tmp)
          .self$workbook$definedNames[toChange] <- tmp
        }
      }

      invisible(.self)
    },

    writeSheetDataXML = function(xldrawingsDir,
      xldrawingsRelsDir,
      xlworksheetsDir,
      xlworksheetsRelsDir) {
      ## write worksheets

      # TODO just seq_along()
      nSheets <- length(.self$worksheets)

      for (i in seq_len(nSheets)) {
        ## Write drawing i (will always exist) skip those that are empty
        if (!identical(.self$drawings[[i]], list())) {
          write_file(
            head = '',
            body = pxml(.self$drawings[[i]]),
            tail = '',
            fl = file.path(xldrawingsDir, stri_join("drawing", i, ".xml"))
          )
          if (!identical(.self$drawings_rels[[i]], list())) {
            write_file(
              head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
              body = pxml(.self$drawings_rels[[i]]),
              tail = "</Relationships>",
              fl = file.path(xldrawingsRelsDir, stri_join("drawing", i, ".xml.rels"))
            )
          }
        } else {
          .self$worksheets[[i]]$drawing <- character()
        }

        ## vml drawing
        if (length(.self$vml_rels[[i]])) {
          file.copy(
            from = .self$vml_rels[[i]],
            to = file.path(
              xldrawingsRelsDir,
              stri_join("vmlDrawing", i, ".vml.rels")
            )
          )
        }

        # outlineLevelRow in SheetformatPr
        if ((length(.self$outlineLevels[[i]])) && (!grepl("outlineLevelRow", .self$worksheets[[i]]$sheetFormatPr))) {
          .self$worksheets[[i]]$sheetFormatPr <- gsub("/>", ' outlineLevelRow="1"/>', .self$worksheets[[i]]$sheetFormatPr)
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
            body = .self$worksheets[[i]]$get_prior_sheet_data(),
            fl = file.path(chartSheetDir, stri_join("sheet", i, ".xml"))
          )

          write_file(
            head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            body = pxml(.self$worksheets_rels[[i]]),
            tail = "</Relationships>",
            fl = file.path(chartSheetRelsDir, sprintf("sheet%s.xml.rels", i))
          )
        } else {
          ## Write worksheets
          ws <- .self$worksheets[[i]]
          hasHL <- length(ws$hyperlinks) > 0

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
          cc_rows <- ws$sheet_data$row_attr$r
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
          # if(!identical(empty_row_attr, character()))
          #   row_attr[[empty_row_attr]] <- list()
          # # restore order
          # ws$sheet_data$row_attr <- row_attr[wanted]

          # message(i, " \n")
          write_worksheet_xml_2(
            prior = prior,
            post = post,
            sheet_data = ws$sheet_data,
            cols_attr = ws$cols_attr,
            R_fileName = file.path(xlworksheetsDir, sprintf("sheet%s.xml", i))
          )

          # # why would I want to erase everything in here?
          # worksheets[[i]]$sheet_data$style_id <- integer()


          ## write worksheet rels
          if (length(.self$worksheets_rels[[i]])) {
            ws_rels <- .self$worksheets_rels[[i]]
            if (hasHL) {
              h_inds <- stri_join(seq_along(.self$worksheets[[i]]$hyperlinks), "h")
              ws_rels <-
                c(ws_rels, unlist(
                  lapply(seq_along(h_inds), function(j) {
                    .self$worksheets[[i]]$hyperlinks[[j]]$to_target_xml(h_inds[j])
                  })
                ))
            }

            ## Check if any tables were deleted - remove these from rels
            if (length(.self$tables)) {
              table_inds <- grep("tables/table[0-9].xml", ws_rels)

              if (length(table_inds)) {
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
                table_nms <- attr(.self$tables, "tableName")[inds]
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
      } ## end of loop through nSheets

      invisible(.self)
    },


    setRowHeights = function(sheet, rows, heights) {
      sheet <- .self$validateSheet(sheet)

      ## remove any conflicting heights
      flag <- names(rowHeights[[sheet]]) %in% rows
      if (any(flag)) {
        .self$rowHeights[[sheet]] <- .self$rowHeights[[sheet]][!flag]
      }

      nms <- c(names(.self$rowHeights[[sheet]]), rows)
      allRowHeights <- unlist(c(.self$rowHeights[[sheet]], heights))
      names(allRowHeights) <- nms

      allRowHeights <-
        allRowHeights[order(as.integer(names(allRowHeights)))]

      .self$rowHeights[[sheet]] <- allRowHeights
      invisible(.self)
    },

    groupColumns = function(sheet) {
      # TODO assert class sheet
      sheet <- .self$validateSheet(sheet)

      hidden <- attr(.self$colOutlineLevels[[sheet]], "hidden", exact = TRUE)
      cols <- names(.self$colOutlineLevels[[sheet]])

      if (!grepl("outlineLevelCol", .self$worksheets[[sheet]]$sheetFormatPr)) {
        .self$worksheets[[sheet]]$sheetFormatPr <- sub("/>", ' outlineLevelCol="1"/>', .self$worksheets[[sheet]]$sheetFormatPr)
      }

      # Check if column is already created (by `setColWidths()` or on import)
      # Note that columns are initiated by `setColWidths` first (see: order of execution in `preSaveCleanUp()`)
      if (any(cols %in% names(.self$worksheets[[sheet]]$cols))) {

        for (i in intersect(cols, names(.self$worksheets[[sheet]]$cols))) {
          outline_hidden <- attr(.self$colOutlineLevels[[sheet]], "hidden")[attr(.self$colOutlineLevels[[sheet]], "names") == i]

          if (grepl("outlineLevel", .self$worksheets[[sheet]]$cols[[i]], perl = TRUE)) {
            .self$worksheets[[sheet]]$cols[[i]] <- sub("((?<=hidden=\")(\\w+)\")", paste0(outline_hidden, "\""), .self$worksheets[[sheet]]$cols[[i]], perl = TRUE)
          } else {
            .self$worksheets[[sheet]]$cols[[i]] <- sub("((?<=hidden=\")(\\w+)\")", paste0(outline_hidden, "\" outlineLevel=\"1\""), .self$worksheets[[sheet]]$cols[[i]], perl = TRUE)
          }
        }

        cols <- cols[!cols %in% names(.self$worksheets[[sheet]]$cols)]
        hidden <- attr(.self$colOutlineLevels[[sheet]], "hidden")[attr(.self$colOutlineLevels[[sheet]], "names") %in% cols]
      }

      if (length(cols)) {
        colNodes <- sprintf('<col min="%s" max="%s" outlineLevel="1" hidden="%s"/>', cols, cols, hidden)
        names(colNodes) <- cols
        .self$worksheets[[sheet]]$cols <- append(.self$worksheets[[sheet]]$cols, colNodes)
      }

      invisible(.self)
    },

    groupRows = function(sheet, rows, hidden, levels) {
      sheet <- .self$validateSheet(sheet)


      flag <- names(.self$outlineLevels[[sheet]]) %in% rows
      if (any(flag)) {
        .self$outlineLevels[[sheet]] <- .self$outlineLevels[[sheet]][!flag]
      }

      nms <- c(names(.self$outlineLevels[[sheet]]), rows)

      allOutlineLevels <- unlist(c(.self$outlineLevels[[sheet]], levels))
      names(allOutlineLevels) <- nms

      existing_hidden <- attr(.self$outlineLevels[[sheet]], "hidden", exact = TRUE)
      all_hidden <- c(existing_hidden, as.character(as.integer(hidden)))

      allOutlineLevels <-
        allOutlineLevels[order(as.integer(names(allOutlineLevels)))]

      .self$outlineLevels[[sheet]] <- allOutlineLevels

      attr(.self$outlineLevels[[sheet]], "hidden") <- as.character(as.integer(all_hidden))


      if (!grepl("outlineLevelRow", .self$worksheets[[sheet]]$sheetFormatPr)) {
        .self$worksheets[[sheet]]$sheetFormatPr <- gsub("/>", ' outlineLevelRow="1"/>', .self$worksheets[[sheet]]$sheetFormatPr)
      }

      invisible(.self)
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

      sheet <- .self$validateSheet(sheet)
      sheetNames <- .self$sheet_names
      nSheets <- length(unlist(sheetNames, use.names = FALSE))
      sheetName <- sheetNames[[sheet]]

      .self$colWidths[[sheet]] <- NULL
      .self$sheet_names <- .self$sheet_names[-sheet]

      ## remove last drawings(sheet).xml from Content_Types
      # TODO replace x[!grepl(x)] with grep(values = TRUE, invert = TRUE)
      .self$Content_Types <-
        .self$Content_Types[!grepl(sprintf("drawing%s.xml", nSheets), .self$Content_Types)]

      ## remove highest sheet
      .self$Content_Types <-
        .self$Content_Types[!grepl(sprintf("sheet%s.xml", nSheets), .self$Content_Types)]

      .self$drawings[[sheet]] <- NULL
      .self$drawings_rels[[sheet]] <- NULL

      .self$vml[[sheet]] <- NULL
      .self$vml_rels[[sheet]] <- NULL

      .self$rowHeights[[sheet]] <- NULL
      .self$colOutlineLevels[[sheet]] <- NULL
      .self$outlineLevels[[sheet]] <- NULL
      .self$comments[[sheet]] <- NULL
      .self$threadComments[[sheet]] <- NULL
      .self$isChartSheet <- .self$isChartSheet[-sheet]

      ## sheetOrder
      # TODO use match()?
      toRemove <- which(.self$sheetOrder == sheet)
      .self$sheetOrder[.self$sheetOrder > sheet] <-
        .self$sheetOrder[.self$sheetOrder > sheet] - 1L
      .self$sheetOrder <- .self$sheetOrder[-toRemove]


      ## remove styleObjects
      if (length(.self$styleObjects)) {
        .self$styleObjects <-
          .self$styleObjects[unlist(lapply(.self$styleObjects, "[[", "sheet"), use.names = FALSE) != sheetName]
      }

      ## Need to remove reference from workbook.xml.rels to pivotCache
      removeRels <- grep("pivotTables", .self$worksheets_rels[[sheet]], value = TRUE)
      if (length(removeRels)) {
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

        fileNo <- grep(toRemove, .self$pivotTables.xml.rels)
        toRemove <-
          stri_join(
            sprintf("(pivotCacheDefinition%s\\.xml)", fileNo),
            sep = " ",
            collapse = "|"
          )

        ## remove reference to file from workbook.xml.res
        # TODO grepl() to grep()
        .self$workbook.xml.rels <-
          .self$workbook.xml.rels[!grepl(toRemove, .self$workbook.xml.rels)]
      }

      ## As above for slicers
      ## Need to remove reference from workbook.xml.rels to pivotCache
      removeRels <- grepl("slicers", .self$worksheets_rels[[sheet]])
      if (any(removeRels)) {
        # TODO !grepl() to grep()
        .self$workbook.xml.rels <-
          .self$workbook.xml.rels[!grepl(sprintf("(slicerCache%s\\.xml)", sheet), .self$workbook.xml.rels)]
      }

      ## wont't remove tables and then won't need to reassign table r:id's but will rename them!
      .self$worksheets[[sheet]] <- NULL
      .self$worksheets_rels[[sheet]] <- NULL

      if (length(.self$tables)) {
        tableSheets <- attr(.self$tables, "sheet")
        tableNames <- attr(.self$tables, "tableName")

        inds <-
          tableSheets %in% sheet &
          !grepl("openxlsx_deleted", attr(.self$tables, "tableName"), fixed = TRUE)
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
        unlist(lapply(.self$workbook$sheets, function(x) {
          regmatches(
            x, regexpr('(?<= name=")[^"]+', x, perl = TRUE)
          )
        }))
      .self$workbook$sheets <- .self$workbook$sheets[!sn %in% sheetName]

      ## Reset rIds
      if (nSheets > 1) {
        for (i in (sheet + 1L):nSheets) {
          .self$workbook$sheets <-
            gsub(stri_join("rId", i),
              stri_join("rId", i - 1L),
              .self$workbook$sheets,
              fixed = TRUE
            )
        }
      } else {
        .self$workbook$sheets <- NULL
      }

      ## Can remove highest sheet
      .self$workbook.xml.rels <-
        .self$workbook.xml.rels[!grepl(sprintf("sheet%s.xml", nSheets), .self$workbook.xml.rels)]

      ## definedNames
      if (length(.self$workbook$definedNames)) {
        belongTo <- getDefinedNamesSheet(.self$workbook$definedNames)
        .self$workbook$definedNames <-
          .self$workbook$definedNames[!belongTo %in% sheetName]
      }

      invisible(.self)
    },

    # TODO possible changes to .self
    addDXFS = function(style) {
      # TODO assert_class(style, "Style")
      dxf <- "<dxf>"
      dxf <- stri_join(dxf, .self$createFontNode(style))
      # fillNode <- NULL

      if (!is.null(style$fill$fillFg) | !is.null(style$fill$fillBg)) {
        dxf <- stri_join(dxf, createFillNode(style))
      }

      # TODO go with length()
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
        # return(which(styles$dxfs == dxf) - 1L)
        return(invisible(.self))
      }

      # dxfId <- length(styles$dxfs)
      .self$styles$dxfs <- c(styles$dxfs, dxf)

      # return(dxfId)
      invisible(.self)
    },

    # TODO rename: setDataValidation?
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
      sheet <- .self$validateSheet(sheet)
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


      # TODO consider switch(type, date = ..., time = ..., )
      if (type == "date") {
        # TODO consider origin <- if () ... else ...
        origin <- 25569L
        # TODO would it be faster to just search each .self$workbook instead of
        # trying to unlist and join everything?
        if (grepl(
          'date1904="1"|date1904="true"',
          stri_join(unlist(.self$workbook), sep = " ", collapse = ""),
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
          stri_join(unlist(.self$workbook), sep = " ", collapse = ""),
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
          .self$worksheets[[sheet]]$dataValidations,
          stri_join(header, stri_join(form, collapse = ""), "</dataValidation>")
        )

      invisible(.self)
    },

    # TODO consider some defaults to logicals
    # TODO rename: setDataValidationList?
    dataValidation_list = function(sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      value,
      allowBlank,
      showInputMsg,
      showErrorMsg) {
      sheet <- .self$validateSheet(sheet)
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
        c(.self$worksheets[[sheet]]$dataValidationsLst, xmlData)

      invisible(.self)
    },

    # TODO consider defaults for logicals
    # TODO rename: setConditionFormatting?  Or addConditionalFormatting
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
      sheet <- .self$validateSheet(sheet)
      sqref <-
        stri_join(getCellRefs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        )), collapse = ":")



      ## Increment priority of conditional formatting rule
      if (length(.self$worksheets[[sheet]]$conditionalFormatting)) {
        for (i in length(.self$worksheets[[sheet]]$conditionalFormatting):1) {
          priority <-
            regmatches(
              .self$worksheets[[sheet]]$conditionalFormatting[[i]],
              regexpr(
                '(?<=priority=")[0-9]+',
                .self$worksheets[[sheet]]$conditionalFormatting[[i]],
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
              .self$worksheets[[sheet]]$conditionalFormatting[[i]],
              fixed = TRUE
            )
        }
      }

      nms <- c(names(.self$worksheets[[sheet]]$conditionalFormatting), sqref)

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
            40000L + length(.self$worksheets[[sheet]]$extLst)
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
            .self$worksheets[[sheet]]$extLst,
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

      # TODO replace append() with just c()
      .self$worksheets[[sheet]]$conditionalFormatting <-
        append(.self$worksheets[[sheet]]$conditionalFormatting, cfRule)

      names(.self$worksheets[[sheet]]$conditionalFormatting) <- nms

      invisible(.self)
    },

    # TODO rename: setMergeCells?  Name "conflicts" with element
    mergeCells = function(sheet, startRow, endRow, startCol, endCol) {
      # TODO assert_class() sheet?
      sheet <- .self$validateSheet(sheetName = sheet)

      sqref <-
        getCellRefs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        ))
      exMerges <-
        regmatches(
          .self$worksheets[[sheet]]$mergeCells,
          regexpr("[A-Z0-9]+:[A-Z0-9]+", .self$worksheets[[sheet]]$mergeCells)
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
          .self$worksheets[[sheet]]$mergeCells,
          sprintf(
            '<mergeCell ref="%s"/>',
            stri_join(sqref,
              collapse = ":", sep =
                " "
            )
          )
        )

      invisible(.self)
    },

    # TODO rename: removeMergeCells
    removeCellMerge = function(sheet, startRow, endRow, startCol, endCol) {
      sheet <- .self$validateSheet(sheet)

      sqref <-
        getCellRefs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        ))
      exMerges <-
        regmatches(
          .self$worksheets[[sheet]]$mergeCells,
          regexpr("[A-Z0-9]+:[A-Z0-9]+", .self$worksheets[[sheet]]$mergeCells)
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
        .self$worksheets[[sheet]]$mergeCells[!mergeIntersections]

      invisible(.self)
    },

    # TODO renam to setFreezePanes?
    freezePanes = function(sheet,
      firstActiveRow = NULL,
      firstActiveCol = NULL,
      firstRow = FALSE,
      firstCol = FALSE) {
      sheet <- .self$validateSheet(sheet)
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
          # return(NULL)
          return(invisible(.self))
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

      invisible(.self)
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

      sheet <- .self$validateSheet(sheet)

      imageType <- regmatches(file, gregexpr("\\.[a-zA-Z]*$", file))
      imageType <- gsub("^\\.", "", imageType)

      imageNo <- length((.self$drawings[[sheet]])) + 1L
      mediaNo <- length(.self$media) + 1L

      startCol <- convertFromExcelRef(startCol)

      ## update Content_Types
      if (!any(grepl(stri_join("image/", imageType), .self$Content_Types))) {
        .self$Content_Types <-
          unique(c(
            sprintf(
              '<Default Extension="%s" ContentType="image/%s"/>',
              imageType,
              imageType
            ),
            .self$Content_Types
          ))
      }

      ## drawings rels (Reference from drawings.xml to image file in media folder)
      .self$drawings_rels[[sheet]] <- c(
        .self$drawings_rels[[sheet]],
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
      .self$media <- append(.self$media, tmp)

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
      .self$drawings[[sheet]] <- c(.self$drawings[[sheet]], drawingsXML)

      invisible(.self)
    },

    # TODO consider name .self$workbook_validate() ?
    preSaveCleanUp = function() {
      ## Steps
      # Order workbook.xml.rels:
      #   sheets -> style -> theme -> sharedStrings -> persons -> tables -> calcChain
      # Assign workbook.xml.rels children rIds, seq_along(workbook.xml.rels)
      # Assign workbook$sheets rIds nSheets
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
          .self$workbook$sheets,
          gregexpr('(?<=r:id="rId)[0-9]+', .self$workbook$sheets, perl = TRUE)
        )))

      nSheets <- length(sheetRIds)
      nExtRefs <- length(.self$externalLinks)
      nPivots <- length(.self$pivotDefinitions)

      ## add a worksheet if none added
      if (nSheets == 0) {
        warning("Workbook does not contain any worksheets. A worksheet will be added.",
          call. = FALSE
        )
        .self$addWorksheet("Sheet 1")
        nSheets <- 1L
      }

      ## get index of each child element for ordering
      # TODO replace which(grepl()) to grep()
      sheetInds <- grep( "(worksheets|chartsheets)/sheet[0-9]+\\.xml", .self$workbook.xml.rels)
      stylesInd <- grep("styles\\.xml", .self$workbook.xml.rels)
      themeInd <- grep("theme/theme[0-9]+.xml", .self$workbook.xml.rels)
      connectionsInd <- grep("connections.xml", .self$workbook.xml.rels)
      extRefInds <- grep("externalLinks/externalLink[0-9]+.xml", .self$workbook.xml.rels)
      sharedStringsInd <- grep("sharedStrings.xml", .self$workbook.xml.rels)
      tableInds <- grep("table[0-9]+.xml", .self$workbook.xml.rels)
      personInds <- grep("person.xml", .self$workbook.xml.rels)


      ## Reordering of workbook.xml.rels
      ## don't want to re-assign rIds for pivot tables or slicer caches
      pivotNode <- grep("pivotCache/pivotCacheDefinition[0-9].xml", .self$workbook.xml.rels, value = TRUE)
      slicerNode <- grep("slicerCache[0-9]+.xml", .self$workbook.xml.rels, value = TRUE)

      ## Reorder children of workbook.xml.rels
      .self$workbook.xml.rels <-
        .self$workbook.xml.rels[c(
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
        unlist(lapply(seq_along(.self$workbook.xml.rels), function(i) {
          gsub('(?<=Relationship Id="rId)[0-9]+',
            i,
            .self$workbook.xml.rels[[i]],
            perl = TRUE
          )
        }))

      .self$workbook.xml.rels <- c(.self$workbook.xml.rels, pivotNode, slicerNode)



      if (!is.null(.self$vbaProject)) {
        .self$workbook.xml.rels <-
          c(
            .self$workbook.xml.rels,
            sprintf(
              '<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>',
              1L + length(.self$workbook.xml.rels)
            )
          )
      }

      ## Reassign rId to workbook sheet elements, (order sheets by sheetId first)
      .self$workbook$sheets <-
        unlist(lapply(seq_along(.self$workbook$sheets), function(i) {
          gsub('(?<= r:id="rId)[0-9]+', i, .self$workbook$sheets[[i]], perl = TRUE)
        }))

      ## re-order worksheets if need to
      if (any(.self$sheetOrder != seq_len(nSheets))) {
        .self$workbook$sheets <- .self$workbook$sheets[sheetOrder]
      }



      ## re-assign tabSelected
      state <- rep.int("visible", nSheets)
      # TODO grepl() to grep()
      state[grepl("hidden", .self$workbook$sheets)] <- "hidden"
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
          .self$worksheets[[visible_sheet_index]]$sheetViews,
          ignore.case = TRUE
        )
      if (nSheets > 1) {
        for (i in setdiff(seq_len(nSheets), visible_sheet_index)) {
          .self$worksheets[[i]]$sheetViews <-
            sub(
              ' tabSelected="(1|true|false|0)"',
              ' tabSelected="0"',
              .self$worksheets[[i]]$sheetViews,
              ignore.case = TRUE
            )
        }
      }





      if (length(.self$workbook$definedNames)) {
        # TODO consider .self$get_sheet_names() which orders the sheet names?
        sheetNames <- .self$sheet_names[.self$sheetOrder]

        belongTo <- getDefinedNamesSheet(.self$workbook$definedNames)

        ## sheetNames is in re-ordered order (order it will be displayed)
        newId <- match(belongTo, sheetNames) - 1L
        oldId <-
          as.numeric(regmatches(
            .self$workbook$definedNames,
            regexpr(
              '(?<= localSheetId=")[0-9]+',
              .self$workbook$definedNames,
              perl = TRUE
            )
          ))

        for (i in seq_along(.self$workbook$definedNames)) {
          if (!is.na(newId[i])) {
            .self$workbook$definedNames[[i]] <-
              gsub(
                sprintf('localSheetId=\"%s\"', oldId[i]),
                sprintf('localSheetId=\"%s\"', newId[i]),
                .self$workbook$definedNames[[i]],
                fixed = TRUE
              )
          }
        }
      }




      ## update workbook r:id to match reordered workbook.xml.rels externalLink element
      if (length(extRefInds)) {
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
      #    rep.int(x = NA_integer_, times = worksheets[[i]]$sheet_data$n_elements)
      #}


      for (x in .self$styleObjects) {
        if (length(x$rows) & length(x$cols)) {
          this.sty <- x$style$copy()

          if (!is.null(this.sty$numFmt)) {
            if (this.sty$numFmt$numFmtId == 9999) {
              this.sty$numFmt$numFmtId <- numFmtIds
              numFmtIds <- numFmtIds + 1L
            }
          }

          ## convert sheet name to index
          ## this creates the XML for styles.XML
          .self$updateStyles(this.sty)
        }
      }


      ## Make sure all rowHeights have rows, if not append them!
      for (i in seq_along(.self$worksheets)) {
        if (length(.self$rowHeights[[i]])) {
          rh <- as.integer(names(.self$rowHeights[[i]]))
          missing_rows <- rh[!rh %in% .self$worksheets[[i]]$sheet_data$rows]
          n <- length(missing_rows)

          if (n > 0) {
            #worksheets[[i]]$sheet_data$style_id <-
            #  c(
            #    worksheets[[i]]$sheet_data$style_id,
            #    rep.int(NA_integer_, times = n)
            #  )

            .self$worksheets[[i]]$sheet_data$rows <-
              c(.self$worksheets[[i]]$sheet_data$rows, missing_rows)
            .self$worksheets[[i]]$sheet_data$cols <-
              c(
                .self$worksheets[[i]]$sheet_data$cols,
                rep.int(NA_integer_, times = n)
              )

            .self$worksheets[[i]]$sheet_data$t <-
              c(.self$worksheets[[i]]$sheet_data$t, rep(NA_integer_, times = n))
            .self$worksheets[[i]]$sheet_data$v <-
              c(
                .self$worksheets[[i]]$sheet_data$v,
                rep(NA_character_, times = n)
              )
            .self$worksheets[[i]]$sheet_data$f <-
              c(
                .self$worksheets[[i]]$sheet_data$f,
                rep(NA_character_, times = n)
              )
            .self$worksheets[[i]]$sheet_data$data_count <-
              .self$worksheets[[i]]$sheet_data$data_count + 1L

            .self$worksheets[[i]]$sheet_data$n_elements <-
              as.integer(length(.self$worksheets[[i]]$sheet_data$rows))
          }
        }

        ## write colwidth and coloutline XML
        if (length(.self$colWidths[[i]])) {
          invisible(.self$setColWidths(i))
        }

        if (length(.self$colOutlineLevels[[i]])) {
          invisible(.self$groupColumns(i))
        }
      }

      invisible(.self)
    },

    addStyle = function(sheet, style, rows, cols, stack) {
      sheet <- .self$sheet_names[[sheet]]

      if (length(.self$styleObjects) == 0) {
        .self$styleObjects <- list(list(
          style = style,
          sheet = sheet,
          rows = rows,
          cols = cols
        ))
      } else if (stack) {
        nStyles <- length(.self$styleObjects)

        ## ********** Assume all styleObjects cells have one a single worksheet **********
        ## Loop through existing styleObjects
        # TODO use seq_along()
        newInds <- seq_along(rows)
        keepStyle <- rep(TRUE, nStyles)
        for (i in seq_len(nStyles)) {
          if (sheet == .self$styleObjects[[i]]$sheet) {
            ## Now check rows and cols intersect
            ## toRemove are the elements that the new style doesn't apply to, we remove these from the style object as it
            ## is copied, merged with the new style and given the new data points

            ex_row_cols <-
              stri_join(.self$styleObjects[[i]]$rows, .self$styleObjects[[i]]$cols, sep = "-")
            new_row_cols <- stri_join(rows, cols, sep = "-")


            ## mergeInds are the intersection of the two styles that will need to merge
            mergeInds <- which(new_row_cols %in% ex_row_cols)

            ## newInds are inds that don't exist in the current - this cumulates until the end to see if any are new
            newInds <- newInds[!newInds %in% mergeInds]


            ## If the new style does not merge
            if (length(mergeInds)) {
              to_remove_from_this_style_object <-
                which(ex_row_cols %in% new_row_cols)

              ## the new style intersects with this styleObjects[[i]], we need to remove the intersecting rows and
              ## columns from styleObjects[[i]]
              if (length(to_remove_from_this_style_object)) {
                ## remove these from style object
                .self$styleObjects[[i]]$rows <-
                  .self$styleObjects[[i]]$rows[-to_remove_from_this_style_object]
                .self$styleObjects[[i]]$cols <-
                  .self$styleObjects[[i]]$cols[-to_remove_from_this_style_object]

                if (length(.self$styleObjects[[i]]$rows) == 0 |
                    length(.self$styleObjects[[i]]$cols) == 0) {
                  keepStyle[i] <-
                    FALSE
                } ## this style applies to no rows or columns anymore
              }

              ## append style object for intersecting cells

              ## we are appending a new style
              keepStyle <-
                c(keepStyle, TRUE) ## keepStyle is used to remove styles that apply to 0 rows OR 0 columns

              ## Merge Style and append to styleObjects
              # TODO replace append() with c()
              .self$styleObjects <-
                append(.self$styleObjects, list(
                  list(
                    style = mergeStyle(.self$styleObjects[[i]]$style, newStyle = style),
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
          .self$styleObjects <- .self$styleObjects[keepStyle]
        }

        ## append style object for non-intersecting cells
        if (length(newInds)) {
          # TODO use c() not append()
          .self$styleObjects <- append(.self$styleObjects, list(list(
            style = style,
            sheet = sheet,
            rows = rows[newInds],
            cols = cols[newInds]
          )))
        }
      } else {
        ## else we are not stacking
        # TODO use c() not append()
        .self$styleObjects <- append(.self$styleObjects, list(list(
          style = style,
          sheet = sheet,
          rows = rows,
          cols = cols
        )))
      } ## End if(length(styleObjects)) else if(stack) {}

      invisible(.self)
    },

    createNamedRegion = function(ref1, ref2, name, sheet, localSheetId = NULL) {
      name <- replaceIllegalCharacters(name)

      if (is.null(localSheetId)) {
        .self$workbook$definedNames <- c(
          .self$workbook$definedNames,
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
          .self$workbook$definedNames,
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

      invisible(.self)
    },

    # returns the new tableName -- basically just lowercase
    validate_table_name = function(tableName) {
      tableName <-
        tolower(tableName) ## Excel forces named regions to lowercase

      # TODO set these to warnings? trim and peplace bad characters with

      # TODO add a strict = getOption("openxlsx.tableName.strict", FALSE)
      # param to force these to allow to stopping
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

      # only place where .self is needed
      if (tableName %in% attr(.self$tables, "tableName")) {
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
      if (length(.self$tables)) {
        tableSheets <- attr(.self$tables, "sheet")
        sheetNo <- .self$validateSheet(sheet)

        to_check <-
          which(tableSheets %in% sheetNo &
              !grepl("openxlsx_deleted", attr(.self$tables, "tableName"), fixed = TRUE))

        if (length(to_check)) {
          ## only look at tables on this sheet

          exTable <- .self$tables[to_check]

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
          # TODO use seq_along()
          for (i in seq_along(exTable)) {
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
      } ## end (length(tables))

      invisible(.self)
    },

    # TODO show() to print() with R6
    show = function() {
      exSheets <- .self$sheet_names
      nSheets <- length(exSheets)
      nImages <- length(.self$media)
      nCharts <- length(.self$charts)
      nStyles <- length(.self$styleObjects)

      exSheets <- replaceXMLEntities(exSheets)
      showText <- "A Workbook object.\n"

      ## worksheets
      if (nSheets > 0) {
        showText <- c(showText, "\nWorksheets:\n")

        # TODO use seq_along()
        sheetTxt <- lapply(seq_len(nSheets), function(i) {
          tmpTxt <- sprintf('Sheet %s: "%s"\n', i, exSheets[[i]])

          if (length(.self$rowHeights[[i]])) {
            tmpTxt <-
              append(
                tmpTxt,
                c(
                  "\n\tCustom row heights (row: height)\n\t",
                  stri_join(
                    sprintf("%s: %s", names(.self$rowHeights[[i]]), round(as.numeric(
                      .self$rowHeights[[i]]
                    ), 2)),
                    collapse = ", ",
                    sep = " "
                  )
                )
              )
          }

          if (length(.self$outlineLevels[[i]])) {
            tmpTxt <-
              append(
                tmpTxt,
                c(
                  "\n\tGrouped rows:\n\t",
                  stri_join(
                    sprintf("%s", names(.self$outlineLevels[[i]])),
                    collapse = ", ",
                    sep = " "
                  )
                )
              )
          }

          if (length(.self$colOutlineLevels[[i]])) {
            tmpTxt <-
              append(
                tmpTxt,
                c(
                  "\n\tGrouped columns:\n\t",
                  stri_join(
                    sprintf("%s", names(.self$colOutlineLevels[[i]])),
                    collapse = ", ",
                    sep = " "
                  )
                )
              )
          }

          if (length(.self$colWidths[[i]])) {
            cols <- names(.self$colWidths[[i]])
            widths <- unname(.self$colWidths[[i]])

            # is width() a list or character vector?
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
            sprintf('Image %s: "%s"\n', seq_len(nImages), .self$media)
          )
      }

      if (nCharts > 0) {
        showText <-
          c(
            showText,
            "\nCharts:\n",
            sprintf('Chart %s: "%s"\n', seq_len(nCharts), .self$charts)
          )
      }

      if (nSheets > 0) {
        showText <-
          c(showText, sprintf(
            "Worksheet write order: %s",
            stri_join(.self$sheetOrder, sep = " ", collapse = ", ")
          ))
      }

      cat(unlist(showText))
      invisible(.self)
    },


    # TODO rename: addConditionalFormatCell
    conditionalFormatCell = function(sheet,
      startRow,
      endRow,
      startCol,
      endCol,
      dxfId,
      formula,
      type) {
      .Deprecated()

      sheet <- .self$validateSheet(sheet)
      sqref <-
        stri_join(getCellRefs(data.frame(
          "x" = c(startRow, endRow),
          "y" = c(startCol, endCol)
        )), collapse = ":")

      ## Increment priority of conditional formatting rule
      if (length((.self$worksheets[[sheet]]$conditionalFormatting))) {
        # TODO use seq_along(); then rev()
        for (i in length(.self$worksheets[[sheet]]$conditionalFormatting):1) {
          .self$worksheets[[sheet]]$conditionalFormatting[[i]] <-
            gsub('(?<=priority=")[0-9]+',
              i + 1L,
              .self$worksheets[[sheet]]$conditionalFormatting[[i]],
              perl = TRUE
            )
        }
      }

      nms <- c(names(.self$worksheets[[sheet]]$conditionalFormatting), sqref)

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
            40000L + length(.self$worksheets[[sheet]]$extLst)
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

      # TODO use c() instead of append()
      .self$worksheets[[sheet]]$conditionalFormatting <-
        append(.self$worksheets[[sheet]]$conditionalFormatting, cfRule)

      names(.self$worksheets[[sheet]]$conditionalFormatting) <- nms

      invisible(.self)
    },

    # TODO This currently does not return .self but does alter .self
    loadStyles = function(stylesXML) {
      ## Build style objects from the styles XML
      styles_XML <- read_xml(stylesXML)

      ## Indexed colours
      vals <- xml_node(styles_XML, "styleSheet", "colors", "indexedColors")
      if (length(vals)) {
        .self$styles$indexedColors <-
          stri_join("<colors>", vals, "</colors>")
      }

      .self$styles$numFmts <- numFmts <- xml_node(styles_XML, "styleSheet", "numFmts", "numFmt")
      # numFmts_attr <- getXMLattr(numFmts, "numFmt")

      ## dxf
      .self$styles$dxfs <- dxf <- xml_node(styles_XML, "styleSheet", "dxfs", "dxf")

      tableStyles <- xml_node(styles_XML, "styleSheet", "tableStyles")
      if (length(tableStyles)) {
        .self$styles$tableStyles <- tableStyles
      }

      # extLst <- getXML2(styles_XML, "styleSheet", "extLst")
      # if (length(extLst)) {
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
      if (length(numFmts)) {
        numFmts <- read_numfmt(numFmts)
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

      .self$styleObjectsList <- styleObjects_tmp
      invisible(.self)
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

      invisible(.self)
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

      invisible(.self)
    },

    # TODO should this be a field?
    getCreators = function() {
      current_creator <-
        stri_match(.self$core, regex = "<dc:creator>(.*?)</dc:creator>")[1, 2]

      current_creator_vec <- as.character(stri_split_fixed(
        str = current_creator,
        pattern = ";",
        simplify = T
      ))

      return(current_creator_vec)
    },

    # TODO rename to setLastModifiedBy() ?
    changeLastModifiedBy = function(LastModifiedBy = NULL) {
      if (!is.null(LastModifiedBy)) {
        current_LastModifiedBy <-
          stri_match(.self$core, regex = "<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>")[1, 2]
        .self$core <-
          stri_replace_all_fixed(
            .self$core,
            pattern = current_LastModifiedBy,
            replacement = LastModifiedBy
          )
      }

      invisible(.self)
    },

    surroundingBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType) {
      sheet <- .self$sheet_names[[.self$validateSheet(sheet)]]
      ## steps
      # get column class
      # get corresponding base style

      for (i in seq_len(nCol)) {
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

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
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

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol
              )
            ))

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
              list(
                "style" = sMid,
                "sheet" = sheet,
                "rows" = (startRow + 1L):(startRow + nRow - 2L), # 2nd -> 2nd to last
                "cols" = rep.int(startCol, nRow - 2L)
              )
            ))

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
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

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
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

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol
              )
            ))

            if (nRow > 2) {
              # TODO use c() not append()
              .self$styleObjects <- append(.self$styleObjects, list(
                list(
                  "style" = sMid,
                  "sheet" = sheet,
                  "rows" = (startRow + 1L):(startRow + nRow - 2L), # 2nd -> 2nd to last
                  "cols" = rep.int(startCol, nRow - 2L)
                )
              ))
            }

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
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

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
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

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol + nCol - 1L
              )
            ))

            if (nRow > 2) {
              # TODO use c() not append()
              .self$styleObjects <- append(.self$styleObjects, list(
                list(
                  "style" = sMid,
                  "sheet" = sheet,
                  "rows" = (startRow + 1L):(startRow + nRow - 2L), # 2nd -> 2nd to last
                  "cols" = rep.int(startCol + nCol - 1L, nRow - 2L)
                )
              ))
            }


            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
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

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
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

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
              list(
                "style" = sTop,
                "sheet" = sheet,
                "rows" = startRow,
                "cols" = startCol + i - 1L
              )
            ))

            ## Middle
            if (specialFormat) {
              # TODO use c() not append()
              .self$styleObjects <- append(.self$styleObjects, list(
                list(
                  "style" = sMid,
                  "sheet" = sheet,
                  "rows" = (startRow + 1L):(startRow + nRow - 2L), # 2nd -> 2nd to last
                  "cols" = rep.int(startCol + i - 1L, nRow - 2L)
                )
              ))
            }

            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
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


      invisible(.self)
    },

    # TODO can nCol be defaulted to length(colClasses)?
    # TODO rename to: setRowBorders?
    rowBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType) {
      sheet <- .self$sheet_names[[.self$validateSheet(sheet)]]
      ## steps
      # get column class
      # get corresponding base style

      for (i in seq_len(nCol)) {
        # TODO use seq_along() with colClasses?
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

        # TODO use c() not append()
        .self$styleObjects <- append(.self$styleObjects, list(
          list(
            "style" = sTop,
            "sheet" = sheet,
            "rows" = (startRow):(startRow + nRow - 1L),
            "cols" = rep(startCol + i - 1L, nRow)
          )
        ))
      } ## End of loop through columns


      invisible(.self)
    },

    # TODO can probably remove nCol?
    # TODO rename to setColumnBorders
    columnBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType) {
      sheet <- .self$sheet_names[[.self$validateSheet(sheet)]]
      ## steps
      # get column class
      # get corresponding base style

      for (i in seq_len(nCol)) {
        tmp <- genBaseColStyle(colClasses[[i]])
        colStyle <- tmp$style
        # TODO Is specialFormat used?
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

          # TODO use c() not append()
          .self$styleObjects <- append(.self$styleObjects, list(
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

          # TODO use c() not append()
          .self$styleObjects <- append(.self$styleObjects, list(
            list(
              "style" = sTop,
              "sheet" = sheet,
              "rows" = startRow,
              "cols" = colInd
            )
          ))

          if (nRow > 2) {
            # TODO use c() not append()
            .self$styleObjects <- append(.self$styleObjects, list(
              list(
                "style" = sMid,
                "sheet" = sheet,
                "rows" = (startRow + 1L):(startRow + nRow - 2L),
                "cols" = rep(colInd, nRow - 2L)
              )
            ))
          }


          # TODO use c() not append()
          .self$styleObjects <- append(.self$styleObjects, list(
            list(
              "style" = sBot,
              "sheet" = sheet,
              "rows" = startRow + nRow - 1L,
              "cols" = colInd
            )
          ))
        }
      } ## End of loop through columns


      invisible(.self)
    },

    # TODO safe to remove nCol?
    # TODO is nCol just length(colClasses?)
    # TODO rename to setAllBorders
    allBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType) {
      sheet <- .self$sheet_names[[.self$validateSheet(sheet)]]
      ## steps
      # get column class
      # get corresponding base style

      for (i in seq_len(nCol)) {
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

        # TODO use c() not append()
        .self$styleObjects <- append(.self$styleObjects, list(
          list(
            "style" = sTop,
            "sheet" = sheet,
            "rows" = (startRow):(startRow + nRow - 1L),
            "cols" = rep(startCol + i - 1L, nRow)
          )
        ))
      } ## End of loop through columns


      invisible(.self)
    },

    # TODO does this actually change anything?
    setColWidths = function(sheet) {
      sheet <- .self$validateSheet(sheet)

      widths <- .self$colWidths[[sheet]]
      hidden <- attr(.self$colWidths[[sheet]], "hidden", exact = TRUE)
      if (length(hidden) != length(widths)) {
        hidden <- rep("0", length(widths))
      }

      cols <- names(.self$colWidths[[sheet]])

      autoColsInds <- widths %in% c("auto", "auto2")
      autoCols <- cols[autoColsInds]

      ## If any not auto
      if (any(!autoColsInds)) {
        widths[!autoColsInds] <- as.numeric(widths[!autoColsInds]) + 0.71
      }

      ## If any auto
      if (length(autoCols)) {

        ## only run if data on worksheet
        if (.self$worksheets[[sheet]]$sheet_data$n_elements == 0) {
          missingAuto <- autoCols
        } else if (all(is.na(.self$worksheets[[sheet]]$sheet_data$v))) {
          missingAuto <- autoCols
        } else {

          ## First thing - get base font max character width
          baseFont <- .self$getBaseFont()
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
          allCharWidths <- rep(baseFontCharWidth, .self$worksheets[[sheet]]$sheet_data$n_elements)
          ######### ----------------------------------------------------------------

          ## get char widths for each style object
          if (length(.self$styleObjects) & any(!is.na(.self$worksheets[[sheet]]$sheet_data$style_id))) {
            # TODO is thisSheetName used?
            thisSheetName <- .self$sheet_names[sheet]

            ## Calc font width for all styles on this worksheet
            styleIds <- .self$worksheets[[sheet]]$sheet_data$style_id
            styObSubet <- .self$styleObjects[sort(unique(styleIds))]
            stySubset <- lapply(styObSubet, "[[", "style")

            ## loop through stlye objects assignin a charWidth else baseFontCharWidth
            styleCharWidths <- sapply(stySubset, get_style_max_char_width, USE.NAMES = FALSE)


            ## Now assign all cells a character width
            allCharWidths <- styleCharWidths[.self$worksheets[[sheet]]$sheet_data$style_id]
            allCharWidths[is.na(allCharWidths)] <- baseFontCharWidth
          }

          ## Now check for columns that are auto2
          auto2Inds <- which(widths %in% "auto2")
          if (length(auto2Inds) & length(.self$worksheets[[sheet]]$mergeCells)) {

            ## get cell merges
            merged_cells <- regmatches(.self$worksheets[[sheet]]$mergeCells, regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells))

            comps <- lapply(merged_cells, function(rectCoords) unlist(strsplit(rectCoords, split = ":")))
            merge_cols <- lapply(comps, convertFromExcelRef)
            merge_cols <- lapply(merge_cols, function(x) x[x %in% cols[auto2Inds]]) ## subset to auto2Inds

            merge_rows <- lapply(comps, function(x) as.numeric(gsub("[A-Z]", "", x, perl = TRUE)))
            merge_rows <- merge_rows[lengths(merge_cols) > 0]
            merge_cols <- merge_cols[lengths(merge_cols) > 0]

            sd <- .self$worksheets[[sheet]]$sheet_data

            if (length(merge_cols)) {
              # TODO use seq_along()
              all_merged_cells <- lapply(seq_along(merge_cols), function(i) {
                expand.grid(
                  "rows" = min(merge_rows[[i]]):max(merge_rows[[i]]),
                  "cols" = min(merge_cols[[i]]):max(merge_cols[[i]])
                )
              })

              all_merged_cells <- do.call("rbind", all_merged_cells)

              ## only want the sheet data in here
              refs <- paste(all_merged_cells[[1]], all_merged_cells[[2]], sep = ",")
              existing_cells <- paste(.self$worksheets[[sheet]]$sheet_data$rows, .self$worksheets[[sheet]]$sheet_data$cols, sep = ",")
              keep <- which(!existing_cells %in% refs & !is.na(.self$worksheets[[sheet]]$sheet_data$v))

              # TODO ad params to SheetData$new() to simplify this
              sd <- SheetData$new()
              sd$cols <- .self$worksheets[[sheet]]$sheet_data$cols[keep]
              sd$t <- .self$worksheets[[sheet]]$sheet_data$t[keep]
              sd$v <- .self$worksheets[[sheet]]$sheet_data$v[keep]
              sd$n_elements <- length(sd$cols)
              allCharWidths <- allCharWidths[keep]
            } else {
              sd <- get_style_max_char_widthworksheets[[sheet]]$sheet_data
            }
          } else {
            sd <- get_style_max_char_widthworksheets[[sheet]]$sheet_data
          }

          ## Now that we have the max character width for the largest font on the page calculate the column widths
          calculatedWidths <- calc_column_widths(
            sheet_data = sd,
            sharedStrings = unlist(.self$sharedStrings, use.names = FALSE),
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

      # TODO remove commented out code?

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
      #if (length(cols)) {
      #  colNodes <- sprintf('<col min="%s" max="%s" width="%s" hidden="%s" customWidth="1"/>', cols, cols, widths, hidden)
      #  names(colNodes) <- cols
      #  worksheets[[sheet]]$cols <- append(worksheets[[sheet]]$cols, colNodes)
      #}

      invisible(.self)
    },

    # TODO add default values?
    writeData = function(df, sheet, startRow, startCol, colNames, colClasses, hlinkNames, keepNA, na.string, list_sep) {
      sheet <- .self$validateSheet(sheet)
      nCols <- ncol(df)
      nRows <- nrow(df)
      df_nms <- names(df)

      allColClasses <- unlist(colClasses)
      df <- as.list(df)

      ######################################################################
      ## standardise all column types


      ## pull out NaN values
      nans <- unlist(lapply(seq_len(nCols), function(i) {
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
        if (grepl('date1904="1"|date1904="true"', stri_join(unlist(.self$workbook), collapse = ""), ignore.case = TRUE)) {
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
        if (length(pInds) & nRows > 0) {
          parseOffset <- function(tz) {
            suppressWarnings(
              ifelse(stri_sub(tz, 1, 1) == "+", 1L, -1L)
              * (as.integer(stri_sub(tz, 2, 3)) + as.integer(stri_sub(tz, 4, 5)) / 60) / 24
            )
          }

          t <- lapply(df[pInds], function(x) format(x, "%z"))
          offSet <- lapply(t, parseOffset)
          offSet <- lapply(offSet, function(x) ifelse(is.na(x), 0, x))

          for (i in seq_along(pInds)) {
            df[[pInds[i]]] <- as.numeric(as.POSIXct(df[[pInds[i]]])) / 86400 + origin + offSet[[i]]
          }
        }
      }

      # TODO for these if () ... for (i in ...); just use the loop?

      ## convert any Dates to integers and create date style object
      if (any(c("currency", "accounting", "percentage", "3", "comma") %in% allColClasses)) {
        cInds <- which(sapply(colClasses, function(x) any(c("accounting", "currency", "percentage", "3", "comma") %in% tolower(x))))
        for (i in cInds) {
          df[[i]] <- as.numeric(gsub("[^0-9\\.-]", "", df[[i]], perl = TRUE))
        }
        class(df[[i]]) <- "numeric"
        # TODO is the above line a typo?  Only assign numeric to the last column?
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

        # TODO use if () ... else ...

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
        t[is.na(v)] <- NA_character_
        v[is.na(v)] <- NA_character_
      }

      ## If any NaN values
      if (length(nans)) {
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
      f_in <- rep.int(NA_character_, length(t))
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
      if (length(hyperlink_cols)) {
        hyperlink_inds <- sort(unlist(lapply(hyperlink_cols, function(i) i + (1:(nRows - colNames) - 1) * nCols + (colNames * nCols)), use.names = FALSE))
        na_hyperlink <- intersect(hyperlink_inds, which(is.na(t)))

        if (length(hyperlink_inds)) {
          t[t %in% 9] <- 1L ## set cell type to "s"

          hyperlink_refs <- convert_to_excel_ref_expand(cols = hyperlink_cols + startCol - 1, LETTERS = LETTERS, rows = as.character((startRow + colNames):(startRow + nRows - 1L)))

          if (length(na_hyperlink)) {
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
          newhl <- lapply(seq_along(hyperlink_inds), function(i) {
            Hyperlink$new(ref = hyperlink_refs[i], target = targets[i], location = NULL, display = NULL, is_external = TRUE)
          })

          .self$worksheets[[sheet]]$hyperlinks <- append(worksheets[[sheet]]$hyperlinks, newhl)
        }
      }


      ## convert all strings to references in sharedStrings and update values (v)
      strFlag <- which(t == 1L)
      newStrs <- v[strFlag]
      if (length(newStrs)) {
        newStrs <- replaceIllegalCharacters(newStrs)
        newStrs <- stri_join("<si><t xml:space=\"preserve\">", newStrs, "</t></si>")

        uNewStr <- unique(newStrs)

        .self$updateSharedStrings(uNewStr)
        v[strFlag] <- match(newStrs, sharedStrings) - 1L
      }

      # ## Create cell list of lists
      .self$worksheets[[sheet]]$sheet_data$write(
        rows_in = startRow:(startRow + nRows - 1L),
        cols_in = startCol:(startCol + nCols - 1L),
        t_in = t,
        v_in = v,
        f_in = f_in,
        any_functions = any_functions
      )



      invisible(.self)
    }
  )

)
