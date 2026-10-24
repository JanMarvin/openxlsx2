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
#' @param sheet sheets (index or name) to convert. Used if `select` is `TRUE`
#'   and `data_only` is `TRUE`. All other sheets are written empty.
#' @param select use `sheet`
#' @return the paths of the files in `xmlDir`
#' @noRd
xls_unpack <- function(file, xmlDir, data_only = FALSE, debug = FALSE, sheet = NULL, select = FALSE) {

  for (dir in c("_rels", "xl/_rels", "xl/worksheets", "xl/worksheets/_rels", "xl/chartsheets", "xl/chartsheets/_rels", "xl/tables", "xl/drawings", "xl/drawings/_rels", "xl/charts", "xl/media", "xl/externalLinks/_rels", "xl/pivotTables/_rels", "xl/pivotCache/_rels", "docProps")) {
    dir.create(file.path(xmlDir, dir), recursive = TRUE)
  }

  xml_decl <- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  path <- function(...) file.path(xmlDir, ...)

  stream <- path("xl", "workbook.xls")
  xls_extract(file, stream, !data_only, debug)

  nsheets <- workbook_xls(stream, path("xl", "workbook.xml"), debug)
  styles_xls(stream, path("xl", "styles.xml"), debug)
  sharedstrings_xls(stream, path("xl", "sharedStrings.xml"), debug)
  has_strings <- any(grepl("<si>", readLines(path("xl", "sharedStrings.xml"), warn = FALSE, encoding = "UTF-8"), fixed = TRUE))
  has_theme <- FALSE
  if (!data_only && theme_xls(stream, path("xl", "theme.zip"), debug) > 0L) {
    exdir <- path("xl", "theme_zip")
    files <- tryCatch(
      utils::unzip(path("xl", "theme.zip"), exdir = exdir),
      error = function(e) character(), warning = function(w) character()
    )
    theme_file <- files[basename(files) == "theme1.xml"]
    # a damaged theme is left out, the workbook then uses the default theme
    theme_ok <- length(theme_file) == 1L && tryCatch({
      read_xml(theme_file, pointer = FALSE)
      all(validUTF8(readLines(theme_file, warn = FALSE, encoding = "UTF-8")))
    }, error = function(e) FALSE, warning = function(w) FALSE)
    if (theme_ok) {
      dir.create(path("xl", "theme"))
      file.copy(theme_file, path("xl", "theme", "theme1.xml"))
      has_theme <- TRUE
    }
    unlink(c(path("xl", "theme.zip"), exdir), recursive = TRUE)
  }
  media <- if (data_only) character() else media_xls(stream, path("xl", "media"), debug)
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
  npivots <- 0L
  cache_ids <- integer()
  keep <- seq_len(nsheets)
  if (data_only && select) {
    wb_head <- paste(readLines(path("xl", "workbook.xml"), warn = FALSE, encoding = "UTF-8"), collapse = "")
    head_tags <- regmatches(wb_head, gregexpr("<sheet [^>]*/>", wb_head))[[1]]
    head_names <- sub(".*name=\"([^\"]*)\".*", "\\1", head_tags)
    head_names <- stringi::stri_replace_all_fixed(head_names, rev(legal_sub()[1:5]), rev(legal_chars()[1:5]), vectorize_all = FALSE)
    keep <- if (is.numeric(sheet)) as.integer(sheet) else match(sheet, head_names)
    keep <- keep[!is.na(keep) & keep >= 1L & keep <= nsheets]
  }
  for (i in seq_len(nsheets)) {
    if (!i %in% keep) {
      writeLines(
        c(xml_decl, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'),
        path("xl", kind[i], sprintf("sheet%d.xml", i))
      )
      next
    }
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
      path("xl"),
      file,
      npivots + 1L,
      cache_ids,
      data_only,
      debug
    )
    if (n[1] > 0L) ncomments <- ncomments + 1L
    ntables <- ntables + n[2]
    ndrawings <- ndrawings + n[3]
    ncharts <- ncharts + n[4]
    npivots <- npivots + n[5]
    cache_ids <- c(cache_ids, n[-(1:5)])
  }
  ncaches <- length(cache_ids)

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
      sprintf('<Override PartName="/xl/pivotTables/pivotTable%d.xml" ContentType="%s.pivotTable+xml"/>', seq_len(npivots), ns_ml),
      sprintf('<Override PartName="/xl/pivotCache/pivotCacheDefinition%d.xml" ContentType="%s.pivotCacheDefinition+xml"/>', seq_len(ncaches), ns_ml),
      sprintf('<Override PartName="/xl/pivotCache/pivotCacheRecords%d.xml" ContentType="%s.pivotCacheRecords+xml"/>', seq_len(ncaches), ns_ml),
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
      sprintf('<Relationship Id="rId%d" Type="%s/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition%d.xml"/>', nsheets + 4L + nextern + seq_len(ncaches), ns_doc, seq_len(ncaches)),
      "</Relationships>"
    ),
    path("xl", "_rels", "workbook.xml.rels"),
    useBytes = TRUE
  )

  if (ncaches > 0L) {
    wbfile <- path("xl", "workbook.xml")
    wbxml <- paste(readLines(wbfile, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
    caches <- sprintf('<pivotCache cacheId="%d" r:id="rId%d"/>', seq_len(ncaches), nsheets + 4L + nextern + seq_len(ncaches))
    wbxml <- sub("</workbook>", paste0("<pivotCaches>", paste(caches, collapse = ""), "</pivotCaches></workbook>"), wbxml, fixed = TRUE)
    writeLines(wbxml, wbfile, useBytes = TRUE)
  }

  if (!data_only) {
    xls_chart_caches(xmlDir)
    xls_filter_names(xmlDir, nsheets, types)
  }
  if (!has_strings) unlink(path("xl", "sharedStrings.xml"))

  # unzip() returns paths with forward slashes on Windows, wb_load() compares
  # these paths with the output of dirname(), which uses forward slashes too
  files <- list.files(xmlDir, recursive = TRUE, full.names = TRUE, all.files = TRUE)
  gsub("\\", "/", files, fixed = TRUE)
}

xls_chart_caches <- function(xmlDir) {
  charts <- list.files(file.path(xmlDir, "xl", "charts"), pattern = "^chart[0-9]+[.]xml$", full.names = TRUE)
  if (!length(charts)) return(invisible(NULL))

  unescape <- function(x) {
    stringi::stri_replace_all_fixed(x, rev(legal_sub()[1:5]), rev(legal_chars()[1:5]), vectorize_all = FALSE)
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

  sheets <- new.env() # nolint
  read_sheet <- function(i) {
    key <- as.character(i)
    if (!is.null(sheets[[key]])) return(sheets[[key]])
    file <- file.path(xmlDir, "xl", "worksheets", sprintf("sheet%d.xml", i))
    cells <- list(ref = character(), num = logical(), text = character())
    if (file.exists(file)) {
      xml <- paste(readLines(file, warn = FALSE, encoding = "UTF-8"), collapse = "")
      m <- regmatches(xml, gregexpr("<c r=\"[A-Z]+[0-9]+\"[^>]*?(/>|>.*?</c>)", xml))[[1]]
      m <- m[grepl("<v>", m, fixed = TRUE)]
      ref <- sub("^<c r=\"([A-Z]+[0-9]+)\".*", "\\1", m)
      open_tag <- sub("^(<c [^>]*>).*", "\\1", m)
      type <- ifelse(grepl(" t=\"", open_tag, fixed = TRUE), sub(".* t=\"([a-zA-Z]+)\".*", "\\1", open_tag), "n")
      val <- sub(".*<v>(.*?)</v>.*", "\\1", m)
      is_s <- type == "s"
      is_str <- type %in% c("str", "inlineStr", "e")
      is_b <- type == "b"
      text <- val
      text[is_s] <- sst[as.integer(val[is_s]) + 1L]
      text[is_str] <- unescape(val[is_str])
      text[is_b] <- ifelse(val[is_b] == "1", "TRUE", "FALSE")
      cells <- list(ref = ref, num = !(is_s | is_str | is_b), text = text)
    }
    sheets[[key]] <- cells
    cells
  }

  fill <- function(f) {
    f <- unescape(f)
    m <- regmatches(f, regexec("^(?:'((?:[^']|'')+)'|([^'!]+))!\\$?([A-Z]+)\\$?([0-9]+)(?::\\$?([A-Z]+)\\$?([0-9]+))?$", f, perl = TRUE))[[1]]
    if (!length(m)) return(NULL)
    name <- if (nzchar(m[2])) gsub("''", "'", m[2], fixed = TRUE) else m[3]
    sheet <- match(name, sheet_names)
    if (is.na(sheet)) return(NULL)
    c1 <- col2int(m[4])
    r1 <- as.integer(m[5])
    c2 <- if (nzchar(m[6])) col2int(m[6]) else c1
    r2 <- if (nzchar(m[7])) as.integer(m[7]) else r1
    cells <- read_sheet(sheet)

    rows <- r1:r2
    cols <- int2col(c1:c2)
    at <- match(paste0(rep(cols, times = length(rows)), rep(rows, each = length(cols))), cells$ref)
    pts <- cells$text[at]
    nums <- !is.na(at) & cells$num[at]
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
      pt_xml <- sprintf("<c:pt idx=\"%d\"><c:v>%s</c:v></c:pt>", idx, replace_legal_chars(got$pts[idx + 1L]))
      new <- if (as_num) {
        sprintf("<c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"%d\"/>%s</c:numCache></c:numRef>",
                replace_legal_chars(unescape(f)), length(got$pts), paste(pt_xml, collapse = ""))
      } else {
        sprintf("<c:strRef><c:f>%s</c:f><c:strCache><c:ptCount val=\"%d\"/>%s</c:strCache></c:strRef>",
                replace_legal_chars(unescape(f)), length(got$pts), paste(pt_xml, collapse = ""))
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
  sheet_names <- stringi::stri_replace_all_fixed(sheet_names, rev(legal_sub()[1:5]), rev(legal_chars()[1:5]), vectorize_all = FALSE)

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
    sheet_ref <- replace_legal_chars(gsub("'", "''", sheet_names[i], fixed = TRUE))
    new_names <- c(new_names, sprintf(
      "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"%d\" hidden=\"1\">'%s'!%s</definedName>",
      i - 1L, sheet_ref, paste(absolute, collapse = ":")
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
