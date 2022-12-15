wb_clone_worksheet_impl <- function(self, private, old, new) {
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
    if (self$isChartSheet[old]) {
      sprintf('<Override PartName="/xl/chartsheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>', newSheetIndex)
    } else {
      sprintf('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', newSheetIndex)
    }
  ))

  ## Update xl/rels
  self$append(
    "workbook.xml.rels",
    if (self$isChartSheet[old]) {
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
  self$isChartSheet[[newSheetIndex]]   <- self$isChartSheet[[old]]
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
}
