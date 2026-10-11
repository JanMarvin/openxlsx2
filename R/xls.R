#' Convert an xls file into the parts of an xlsx file
#'
#' Like for xlsb files, the xls file is converted into pseudo xlsx parts, that
#' are read by the usual xlsx code. In contrast to xlsx and xlsb an xls file is
#' not a zip file with a part for each of workbook, styles, shared strings and
#' sheets. It has a single stream of records (BIFF8). The stream is saved and then
#' read once for each part.
#'
#' @param file path to the xls file
#' @param xmlDir temporary directory the parts are written to
#' @param data_only skip everything that is not needed to read cell values
#' @param debug print the records while reading
#' @return the paths of the files in `xmlDir`
#' @noRd
xls_unpack <- function(file, xmlDir, data_only = FALSE, debug = FALSE) {

  for (dir in c("_rels", "xl/_rels", "xl/worksheets", "xl/worksheets/_rels", "xl/chartsheets", "xl/chartsheets/_rels", "xl/tables", "xl/drawings", "xl/drawings/_rels", "xl/charts", "xl/media", "xl/externalLinks/_rels", "docProps")) {
    dir.create(file.path(xmlDir, dir), recursive = TRUE)
  }

  xml_decl <- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  path <- function(...) file.path(xmlDir, ...)

  stream <- path("xl", "workbook.xls")
  xls_extract(file, stream, debug)

  nsheets <- workbook_xls(stream, path("xl", "workbook.xml"), debug)
  styles_xls(stream, path("xl", "styles.xml"), debug)
  sharedstrings_xls(stream, path("xl", "sharedStrings.xml"), debug)
  has_strings <- any(grepl("<si>", readLines(path("xl", "sharedStrings.xml"), warn = FALSE, encoding = "UTF-8"), fixed = TRUE))
  has_theme <- FALSE
  if (!data_only && theme_xls(stream, path("xl", "theme.zip"), debug) > 0L) {
    exdir <- path("xl", "theme_zip")
    files <- utils::unzip(path("xl", "theme.zip"), exdir = exdir)
    theme_file <- files[basename(files) == "theme1.xml"]
    if (length(theme_file) == 1L) {
      dir.create(path("xl", "theme"))
      file.copy(theme_file, path("xl", "theme", "theme1.xml"))
      has_theme <- TRUE
    }
    unlink(c(path("xl", "theme.zip"), exdir), recursive = TRUE)
  }
  media <- if (data_only) character() else media_xls(stream, path("xl", "media"), debug)
  usize <- attr(media, "usize")
  for (i in which(endsWith(media, ".zlib"))) {
    packed <- path("xl", "media", media[i])
    media[i] <- sub("\\.zlib$", "", media[i])
    zlib <- readBin(packed, "raw", file.size(packed))
    data <- NULL
    if (usize[i] > 0 && usize[i] <= 2^28 && length(zlib) > 6L) {
      # the zlib stream is read as gzip with a bounded size: a file that claims
      # less than it holds can not fill the memory
      gz <- c(as.raw(c(0x1f, 0x8b, 8, 0, 0, 0, 0, 0, 0, 255)), zlib[-c(1:2, length(zlib) - 3:0)], raw(8))
      # the check sum of the made up gzip trailer fails and is reported on stderr
      quiet <- file(nullfile(), "w")
      sink(quiet, type = "message")
      con <- gzcon(rawConnection(gz))
      data <- tryCatch(readBin(con, "raw", usize[i]), error = function(e) NULL)
      close(con)
      sink(type = "message")
      close(quiet)
      if (length(data) != usize[i]) data <- NULL
    }
    if (is.null(data)) media[i] <- "" else writeBin(data, path("xl", "media", media[i]))
    unlink(packed)
  }
  props <- if (data_only) 0L else docprops_xls(file, path("docProps", "core.xml"), path("docProps", "app.xml"))
  has_core <- bitwAnd(props, 1L) > 0L
  has_app <- bitwAnd(props, 2L) > 0L
  nextern <- externallinks_xls(stream, path("xl", "externalLinks"), debug)

  types <- sheettypes_xls(stream)
  kind <- ifelse(types == 0L, "worksheets", "chartsheets")

  # comments and tables are numbered in the order of the sheets they are on
  ncomments <- 0L
  ntables <- 0L
  ndrawings <- 0L
  ncharts <- 0L
  for (i in seq_len(nsheets)) {
    n <- worksheet_xls(
      stream, i,
      path("xl", kind[i], sprintf("sheet%d.xml", i)),
      path("xl", kind[i], "_rels", sprintf("sheet%d.xml.rels", i)),
      path("xl", sprintf("comments%d.xml", ncomments + 1L)),
      path("xl", "drawings", sprintf("vmlDrawing%d.vml", ncomments + 1L)),
      ncomments + 1L,
      path("xl", "tables"),
      ntables + 1L,
      media,
      path("xl", "drawings", sprintf("drawing%d.xml", ndrawings + 1L)),
      path("xl", "drawings", "_rels", sprintf("drawing%d.xml.rels", ndrawings + 1L)),
      ndrawings + 1L,
      path("xl", "charts"),
      ncharts + 1L,
      data_only,
      debug
    )
    if (n[1] > 0L) ncomments <- ncomments + 1L
    ntables <- ntables + n[2]
    ndrawings <- ndrawings + n[3]
    ncharts <- ncharts + n[4]
  }

  # the parts that come with a zip file
  ns_pkg <- "http://schemas.openxmlformats.org/package/2006/relationships"
  ns_doc <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  ns_ml <- "application/vnd.openxmlformats-officedocument.spreadsheetml"

  writeLines(
    c(
      xml_decl,
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
      '<Default Extension="png" ContentType="image/png"/>',
      '<Default Extension="jpeg" ContentType="image/jpeg"/>',
      '<Default Extension="bmp" ContentType="image/bmp"/>',
      '<Default Extension="emf" ContentType="image/x-emf"/>',
      '<Default Extension="wmf" ContentType="image/x-wmf"/>',
      '<Default Extension="tiff" ContentType="image/tiff"/>',
      sprintf('<Override PartName="/xl/workbook.xml" ContentType="%s.sheet.main+xml"/>', ns_ml),
      if (has_core) '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
      if (has_app) '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
      sprintf('<Override PartName="/xl/styles.xml" ContentType="%s.styles+xml"/>', ns_ml),
      if (has_strings) sprintf('<Override PartName="/xl/sharedStrings.xml" ContentType="%s.sharedStrings+xml"/>', ns_ml),
      sprintf('<Override PartName="/xl/%s/sheet%d.xml" ContentType="%s.%s+xml"/>', kind, seq_len(nsheets), ns_ml, ifelse(types == 0L, "worksheet", "chartsheet")),
      sprintf('<Override PartName="/xl/comments%d.xml" ContentType="%s.comments+xml"/>', seq_len(ncomments), ns_ml),
      if (has_theme) sprintf('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'),
      sprintf('<Override PartName="/xl/externalLinks/externalLink%d.xml" ContentType="%s.externalLink+xml"/>', seq_len(nextern), ns_ml),
      sprintf('<Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', seq_len(ndrawings)),
      sprintf('<Override PartName="/xl/charts/chart%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', seq_len(ncharts)),
      sprintf('<Override PartName="/xl/tables/table%d.xml" ContentType="%s.table+xml"/>', seq_len(ntables), ns_ml),
      "</Types>"
    ),
    path("[Content_Types].xml"),
    useBytes = TRUE
  )

  writeLines(
    c(
      xml_decl,
      sprintf('<Relationships xmlns="%s">', ns_pkg),
      sprintf('<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>', ns_doc),
      if (has_core) '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
      if (has_app) sprintf('<Relationship Id="rId3" Type="%s/extended-properties" Target="docProps/app.xml"/>', ns_doc),
      "</Relationships>"
    ),
    path("_rels", ".rels"),
    useBytes = TRUE
  )

  writeLines(
    c(
      xml_decl,
      sprintf('<Relationships xmlns="%s">', ns_pkg),
      sprintf('<Relationship Id="rId%d" Type="%s/%s" Target="%s/sheet%d.xml"/>', seq_len(nsheets), ns_doc, ifelse(types == 0L, "worksheet", "chartsheet"), kind, seq_len(nsheets)),
      sprintf('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>', nsheets + 1L, ns_doc),
      if (has_strings) sprintf('<Relationship Id="rId%d" Type="%s/sharedStrings" Target="sharedStrings.xml"/>', nsheets + 2L, ns_doc),
      sprintf('<Relationship Id="rId%d" Type="%s/externalLink" Target="externalLinks/externalLink%d.xml"/>', nsheets + 2L + seq_len(nextern), ns_doc, seq_len(nextern)),
      if (has_theme) sprintf('<Relationship Id="rId%d" Type="%s/theme" Target="theme/theme1.xml"/>', nsheets + 3L + nextern, ns_doc),
      "</Relationships>"
    ),
    path("xl", "_rels", "workbook.xml.rels"),
    useBytes = TRUE
  )

  if (!data_only) {
    xls_chart_caches(xmlDir)
    xls_filter_names(xmlDir, nsheets, types)
  }
  if (!has_strings) unlink(path("xl", "sharedStrings.xml"))

  list.files(xmlDir, recursive = TRUE, full.names = TRUE, all.files = TRUE)
}

xls_chart_caches <- function(xmlDir) {
  charts <- list.files(file.path(xmlDir, "xl", "charts"), pattern = "^chart[0-9]+[.]xml$", full.names = TRUE)
  if (!length(charts)) return(invisible(NULL))

  unescape <- function(x) {
    x <- gsub("&lt;", "<", x, fixed = TRUE)
    x <- gsub("&gt;", ">", x, fixed = TRUE)
    x <- gsub("&quot;", "\"", x, fixed = TRUE)
    x <- gsub("&apos;", "'", x, fixed = TRUE)
    gsub("&amp;", "&", x, fixed = TRUE)
  }
  escape <- function(x) {
    x <- gsub("&", "&amp;", x, fixed = TRUE)
    x <- gsub("<", "&lt;", x, fixed = TRUE)
    gsub(">", "&gt;", x, fixed = TRUE)
  }

  wb_xml <- paste(readLines(file.path(xmlDir, "xl", "workbook.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  sheet_tags <- regmatches(wb_xml, gregexpr("<sheet [^>]*/>", wb_xml))[[1]]
  sheet_names <- unescape(sub(".*name=\"([^\"]*)\".*", "\\1", sheet_tags))

  sst_xml <- paste(readLines(file.path(xmlDir, "xl", "sharedStrings.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
  items <- regmatches(sst_xml, gregexpr("<si>.*?</si>|<si/>", sst_xml))[[1]]
  sst <- vapply(items, function(si) {
    txt <- regmatches(si, gregexpr("<t[^>]*>[^<]*</t>", si))[[1]]
    unescape(paste(sub("</t>$", "", sub("^<t[^>]*>", "", txt)), collapse = ""))
  }, character(1), USE.NAMES = FALSE)

  col_number <- function(x) {
    ch <- utf8ToInt(x) - 64L
    Reduce(function(a, b) a * 26L + b, ch, 0L)
  }

  sheets <- list()
  read_sheet <- function(i) {
    key <- as.character(i)
    if (!is.null(sheets[[key]])) return(sheets[[key]])
    file <- file.path(xmlDir, "xl", "worksheets", sprintf("sheet%d.xml", i))
    cells <- list()
    if (file.exists(file)) {
      xml <- paste(readLines(file, warn = FALSE, encoding = "UTF-8"), collapse = "")
      m <- regmatches(xml, gregexpr("<c r=\"[A-Z]+[0-9]+\"[^>]*?(/>|>.*?</c>)", xml))[[1]]
      for (cell in m) {
        ref <- sub("^<c r=\"([A-Z]+[0-9]+)\".*", "\\1", cell)
        open_tag <- sub("^(<c [^>]*>).*", "\\1", cell)
        type <- if (grepl(" t=\"", open_tag, fixed = TRUE)) sub(".* t=\"([a-zA-Z]+)\".*", "\\1", open_tag) else "n"
        val <- if (grepl("<v>", cell, fixed = TRUE)) sub(".*<v>(.*?)</v>.*", "\\1", cell) else NA_character_
        if (is.na(val)) next
        if (type == "s") {
          cells[[ref]] <- list(num = FALSE, text = sst[as.integer(val) + 1L])
        } else if (type %in% c("str", "inlineStr", "e")) {
          cells[[ref]] <- list(num = FALSE, text = unescape(val))
        } else if (type == "b") {
          cells[[ref]] <- list(num = FALSE, text = if (val == "1") "TRUE" else "FALSE")
        } else {
          cells[[ref]] <- list(num = TRUE, text = val)
        }
      }
    }
    sheets[[key]] <<- cells
    cells
  }

  fill <- function(f) {
    f <- unescape(f)
    m <- regmatches(f, regexec("^(?:'((?:[^']|'')+)'|([^'!]+))!\\$?([A-Z]+)\\$?([0-9]+)(?::\\$?([A-Z]+)\\$?([0-9]+))?$", f, perl = TRUE))[[1]]
    if (!length(m)) return(NULL)
    name <- if (nzchar(m[2])) gsub("''", "'", m[2], fixed = TRUE) else m[3]
    sheet <- match(name, sheet_names)
    if (is.na(sheet)) return(NULL)
    c1 <- col_number(m[4]); r1 <- as.integer(m[5])
    c2 <- if (nzchar(m[6])) col_number(m[6]) else c1
    r2 <- if (nzchar(m[7])) as.integer(m[7]) else r1
    cells <- read_sheet(sheet)

    pts <- character()
    nums <- logical()
    for (r in r1:r2) {
      for (cc in c1:c2) {
        letters_ <- character()
        n <- cc
        while (n > 0L) {
          letters_ <- c(LETTERS[(n - 1L) %% 26L + 1L], letters_)
          n <- (n - 1L) %/% 26L
        }
        cell <- cells[[paste0(paste(letters_, collapse = ""), r)]]
        pts <- c(pts, if (is.null(cell)) NA_character_ else cell$text)
        nums <- c(nums, !is.null(cell) && cell$num)
      }
    }
    list(pts = pts, numeric = all(nums[!is.na(pts)]))
  }

  for (chart in charts) {
    xml <- paste(readLines(chart, warn = FALSE, encoding = "UTF-8"), collapse = "")
    refs <- regmatches(xml, gregexpr("<c:(?:num|str)Ref><c:f>[^<]*</c:f></c:(?:num|str)Ref>", xml, perl = TRUE))[[1]]
    for (ref in unique(refs)) {
      f <- sub("^<c:(?:num|str)Ref><c:f>([^<]*)</c:f>.*", "\\1", ref, perl = TRUE)
      was_string <- startsWith(ref, "<c:strRef>")
      got <- fill(f)
      if (is.null(got)) next
      idx <- which(!is.na(got$pts)) - 1L
      as_num <- got$numeric && !was_string
      pt_xml <- sprintf("<c:pt idx=\"%d\"><c:v>%s</c:v></c:pt>", idx, escape(got$pts[idx + 1L]))
      new <- if (as_num) {
        sprintf("<c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"%d\"/>%s</c:numCache></c:numRef>",
                escape(unescape(f)), length(got$pts), paste(pt_xml, collapse = ""))
      } else {
        sprintf("<c:strRef><c:f>%s</c:f><c:strCache><c:ptCount val=\"%d\"/>%s</c:strCache></c:strRef>",
                escape(unescape(f)), length(got$pts), paste(pt_xml, collapse = ""))
      }
      xml <- gsub(ref, new, xml, fixed = TRUE)
    }
    writeLines(xml, chart, useBytes = TRUE)
  }
  invisible(NULL)
}

xls_filter_names <- function(xmlDir, nsheets, types) {
  wb_file <- file.path(xmlDir, "xl", "workbook.xml")
  wb_xml <- paste(readLines(wb_file, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  sheet_tags <- regmatches(wb_xml, gregexpr("<sheet [^>]*/>", wb_xml))[[1]]
  sheet_names <- sub(".*name=\"([^\"]*)\".*", "\\1", sheet_tags)

  new_names <- character()
  for (i in seq_len(nsheets)) {
    if (types[i] != 0L) next
    sheet_xml <- paste(readLines(file.path(xmlDir, "xl", "worksheets", sprintf("sheet%d.xml", i)), warn = FALSE, encoding = "UTF-8"), collapse = "")
    ref <- regmatches(sheet_xml, regexec("<autoFilter ref=\"([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)\"", sheet_xml))[[1]]
    if (length(ref) < 2L) next
    has_name <- grepl(sprintf("<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"%d\"", i - 1L), wb_xml, fixed = TRUE)
    if (has_name) next
    cells <- strsplit(ref[2], ":", fixed = TRUE)[[1]]
    absolute <- sub("^([A-Z]+)([0-9]+)$", "$\\1$\\2", cells)
    new_names <- c(new_names, sprintf(
      "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"%d\" hidden=\"1\">'%s'!%s</definedName>",
      i - 1L, gsub("'", "''", sheet_names[i], fixed = TRUE), paste(absolute, collapse = ":")
    ))
  }
  if (!length(new_names)) return(invisible(NULL))

  if (grepl("</definedNames>", wb_xml, fixed = TRUE)) {
    wb_xml <- sub("</definedNames>", paste0(paste(new_names, collapse = ""), "</definedNames>"), wb_xml, fixed = TRUE)
  } else if (grepl("<calcPr", wb_xml, fixed = TRUE)) {
    wb_xml <- sub("<calcPr", paste0("<definedNames>", paste(new_names, collapse = ""), "</definedNames><calcPr"), wb_xml, fixed = TRUE)
  } else {
    wb_xml <- sub("</workbook>", paste0("<definedNames>", paste(new_names, collapse = ""), "</definedNames></workbook>"), wb_xml, fixed = TRUE)
  }
  writeLines(wb_xml, wb_file, useBytes = TRUE)
  invisible(NULL)
}
