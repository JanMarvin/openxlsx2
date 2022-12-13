wb_add_image_impl <- function(
    self,
    private,
    sheet     = current_sheet(),
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
  width <- as.integer(round(width * 914400L, 0)) # (EMUs per inch)
  height <- as.integer(round(height * 914400L, 0)) # (EMUs per inch)

  ## within the sheet the drawing node's Id refernce an id in the sheetRels
  ## sheet rels reference the drawingi.xml file
  ## drawingi.xml refernece drawingRels
  ## drawing rels reference an image in the media folder
  ## worksheetRels(sheet(i)) references drawings(j)

  sheet <- private$get_sheet_index(sheet)

  # TODO tools::file_ext() ...
  imageType <- regmatches(file, gregexpr("\\.[a-zA-Z]*$", file))
  imageType <- gsub("^\\.", "", imageType)

  drawing_sheet <- 1
  if (length(self$worksheets_rels[[sheet]])) {
    relship <- rbindlist(xml_attr(self$worksheets_rels[[sheet]], "Relationship"))
    relship$typ <- basename(relship$Type)
    drawing_sheet  <- as.integer(gsub("\\D+", "", relship$Target[relship$typ == "drawing"]))
  }

  drawing_len <- 0
  if (!all(self$drawings_rels[[drawing_sheet]] == ""))
    drawing_len <- length(xml_node(unlist(self$drawings_rels[[drawing_sheet]]), "Relationship"))

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

  # update worksheets_rels
  if (length(self$worksheets_rels[[sheet]]) == 0) {
    self$worksheets_rels[[sheet]] <- sprintf('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing%s.xml"/>', imageNo) ## will always be 1
  }

  # update drawings_rels
  old_drawings_rels <- unlist(self$drawings_rels[[drawing_sheet]])
  if (all(old_drawings_rels == "")) old_drawings_rels <- NULL
  ## drawings rels (Reference from drawings.xml to image file in media folder)
  self$drawings_rels[[drawing_sheet]] <- c(
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
  if (all(self$drawings[[drawing_sheet]] == "")) {
    xml_attr <- c(
      "xmlns:xdr" = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      "xmlns:a" = "http://schemas.openxmlformats.org/drawingml/2006/main",
      "xmlns:r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    self$drawings[[drawing_sheet]] <- xml_node_create("xdr:wsDr", xml_children = drawingsXML, xml_attributes = xml_attr)
  } else {
    self$drawings[[drawing_sheet]] <- xml_add_child(self$drawings[[drawing_sheet]], drawingsXML)
  }

  # Finally we must assign the drawing to the sheet, if no drawing is assigned. If drawing is not
  # empty, we must assume that the rId matches another drawing rId in worksheets_rels
  if (identical(self$worksheets[[sheet]]$drawing, character()))
    self$worksheets[[sheet]]$drawing <- '<drawing r:id=\"rId1\"/>' ## will always be 1

  invisible(self)
}

wb_add_plot_impl <- function(
    self,
    private,
    sheet     = current_sheet(),
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
}

wb_add_drawing_impl <- function(
    self,
    private,
    sheet = current_sheet(),
    xml,
    dims  = "A1:H8"
) {
  sheet <- private$get_sheet_index(sheet)

  xml <- read_xml(xml, pointer = FALSE)

  if (!(xml_node_name(xml) == "xdr:wsDr")) {
    error("xml needs to be a drawing.")
  }

  grpSp <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:grpSp")
  ext   <- xml_node(xml, "xdr:wsDr", "xdr:absoluteAnchor", "xdr:ext")

  # include rvg graphic from specific position to one or two cell anchor
  if (!is.null(dims) && xml_node_name(xml, "xdr:wsDr") == "xdr:absoluteAnchor") {

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
        "<xdr:clientData/>"
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
  drawings <- self$drawings[[sheet]]
  if (drawings == "") {
    drawings <- xml
  } else {
    drawing_type <- xml_node_name(xml, "xdr:wsDr")
    xml_drawing <- xml_node(xml, "xdr:wsDr", drawing_type)
    drawings <- xml_add_child(drawings, xml_drawing)
  }
  self$drawings[[sheet]] <- drawings

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
    self$worksheets_rels[[sheet]] <- sprintf("<Relationship Id=\"rId%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing%s.xml\"/>", next_relship, sheet)
    self$worksheets[[sheet]]$drawing <- sprintf("<drawing r:id=\"rId%s\"/>", next_relship)
  }

  invisible(self)
}

wb_add_chart_xml_impl <- function(
    self,
    private,
    sheet   = current_sheet(),
    xml,
    dims    = "A1:H8"
) {
  dims_list <- strsplit(dims, ":")[[1]]
  cols <- col2int(dims_list)
  rows <- as.numeric(gsub("\\D+", "", dims_list))

  sheet <- private$get_sheet_index(sheet)

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

  len_drawing <- length(xml_node_name(self$drawings[[sheet]], "xdr:wsDr")) + 1L

  from <- c(cols[1] - 1L, rows[1] - 1L)
  to   <- c(cols[2], rows[2])

  # create drawing. add it to self$drawings, the worksheet and rels
  self$add_drawing(
    sheet = sheet,
    xml = drawings(len_drawing, from, to),
    dims = dims
  )

  self$drawings_rels[[sheet]] <- drawings_rels(self$drawings_rels[[sheet]], next_chart)
  invisible(self)
}

wb_add_mschart_impl <- function(
    self,
    private,
    sheet   = current_sheet(),
    dims    = "B2:H8",
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
    self$add_chart_xml(sheet = sheet, xml = out_xml, dims = dims)
  } else {
    self$add_data(
      x = graph$data_series
    )$add_chart_xml(
      sheet = sheet,
      xml = out_xml,
      dims = dims
    )
  }
}
