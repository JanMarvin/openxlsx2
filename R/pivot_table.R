# unique is not enough! gh#713
distinct <- function(x) {
  unis <- stringi::stri_unique(x)
  lwrs <- tolower(unis)
  dups <- duplicated(lwrs)
  unis[!dups]
}

# append number of duplicated value.
# @details c("Foo", "foo") -> c("Foo", "foo2")
# @param x a character vector
fix_pt_names <- function(x) {
  lwrs <- tolower(x)

  dups <- duplicated(lwrs)
  dups <- stringi::stri_unique(lwrs[dups])

  for (dup in dups) {
    sel <- which(lwrs == dup)
    for (i in seq_along(sel)[-1]) {
      x[sel[i]] <- paste0(x[sel[i]], i)
    }
  }

  x
}

cacheFields <- function(wbdata, filter, rows, cols, data, slicer, timeline) {
  sapply(
    names(wbdata),
    function(x) {

      dat <- wbdata[[x]]

      vars <- c(filter, rows, cols, data, slicer, timeline)

      is_vars <- x %in% vars
      is_data <- x %in% data &&
        # there is an exception to every rule in pivot tables ...
        # if a data variable is col or row we need items
        !x %in% cols && !x %in% rows && !x %in% filter
      is_char <- is.character(dat)
      is_date <- inherits(dat, "Date") || inherits(dat, "POSIXct")
      if (is_date) {
        dat <- format(dat, format = "%Y-%m-%dT%H:%M:%S")
      }

      is_formula <- inherits(wbdata[[x]], what = "is_formula")

      if (is_vars && !is_data && !is_formula) {
        sharedItem <- vapply(
          distinct(dat),
          function(uni) {
            if (is.na(uni)) {
              xml_node_create("m")
            } else {

              if (is_char) {
                xml_node_create("s", xml_attributes = c(v = uni), escapes = TRUE)
              } else if (is_date) {
                xml_node_create("d", xml_attributes = c(v = uni))
              } else {
                xml_node_create("n", xml_attributes = c(v = uni))
              }
            }

          },
          FUN.VALUE = ""
        )
      } else {
        sharedItem <- NULL
      }

      if (anyNA(dat)) {
        containsBlank <- "1"
        containsSemiMixedTypes <- NULL
      } else {
        containsBlank <- NULL
        # indicates that cell has no blank values
        containsSemiMixedTypes <- "0"
      }

      if (length(sharedItem)) {
        count <- as.character(length(sharedItem))
      } else {
        count <- NULL
      }

      if (is_char) {
        attr <- c(
          containsBlank = containsBlank,
          count = count
        )
      } else if (is_date) {
        attr <- c(
          containsNonDate = "0",
          containsDate = "1",
          # enabled, check side effect
          containsString = "0",
          containsSemiMixedTypes = containsSemiMixedTypes,
          containsBlank = containsBlank,
          # containsMixedTypes = "1",
          minDate = format(min(dat, na.rm = TRUE), format = "%Y-%m-%dT%H:%M:%S"),
          maxDate = format(max(dat, na.rm = TRUE), format = "%Y-%m-%dT%H:%M:%S"),
          count = count
        )

      } else {

        # numeric or integer
        is_int <- all(dat[!is.na(dat)] %% 1 == 0)
        if (is_int) {
          containsInteger <- as_binary(is_int)
        } else {
          containsInteger <- NULL
        }

        attr <- c(
          containsSemiMixedTypes = containsSemiMixedTypes,
          containsString = "0",
          containsBlank = containsBlank,
          containsNumber = "1",
          containsInteger = containsInteger, # double or int?
          minValue = as.character(min(dat, na.rm = TRUE)),
          maxValue = as.character(max(dat, na.rm = TRUE)),
          count = count
        )
      }

      sharedItems <- xml_node_create(
        "sharedItems",
        xml_attributes = attr,
        xml_children = sharedItem
      )

      formula <- NULL
      databaseField <- NULL

      if (is_formula) {
        formula       <- gsub("^=", "", names(vars[vars == x]))
        databaseField <- "0"
        sharedItems   <- NULL
      }

      xml_node_create(
        "cacheField",
        xml_attributes = c(name = x, numFmtId = "0", formula = formula, databaseField = databaseField),
        xml_children = sharedItems
      )
    },
    USE.NAMES = FALSE
  )
}

pivot_def_xml <- function(wbdata, filter, rows, cols, data, slicer, timeline, pcid) {

  ref   <- dataframe_to_dims(attr(wbdata, "dims"))
  sheet <- attr(wbdata, "sheet")
  count <- ncol(wbdata)

  paste0(
    sprintf('<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" mc:Ignorable="xr" r:id="rId1" refreshedBy="openxlsx2" invalid="1" refreshOnLoad="1" refreshedDate="1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="%s">', nrow(wbdata)),
    '<cacheSource type="worksheet"><worksheetSource ref="', ref, '" sheet="', sheet, '"/></cacheSource>',
    '<cacheFields count="', count, '">',
    paste0(cacheFields(wbdata, filter, rows, cols, data, slicer, timeline), collapse = ""),
    '</cacheFields>',
    '<extLst>',
    '<ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}">',
    sprintf('<x14:pivotCacheDefinition pivotCacheId="%s"/>', pcid),
    '</ext>',
    '</extLst>',
    '</pivotCacheDefinition>'
  )
}


pivot_rec_xml <- function(data) {
  paste0(
    sprintf('<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" count="%s">', nrow(data)),
    '</pivotCacheRecords>'
  )
}

pivot_def_rel <- function(n) sprintf("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\" Target=\"pivotCacheRecords%s.xml\"/></Relationships>", n)

pivot_xml_rels <- function(n) sprintf("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition%s.xml\"/></Relationships>", n)

get_items <- function(data, x, item_order, slicer = FALSE, choose = NULL, has_default = TRUE) {
  x <- abs(x)

  dat <- distinct(data[[x]])

  default <- NULL
  if (has_default) default <- "default"

  # check length, otherwise a certain spreadsheet software simply dies
  if (!is.null(item_order)) {

    if (length(item_order) > length(dat)) {
      msg <- sprintf(
        "Length of sort order for '%s' does not match required length. Is %s, needs %s.\nCheck `openxlsx2:::distinct()` for the correct length. Resetting.",
        names(data[x]), length(item_order), length(dat)
      )
      warning(msg)
      item_order <- NULL
    }

    if (is.character(item_order)) {
      # add remaining items
      if (length(item_order) < length(dat)) {
        item_order <- c(item_order, dat[!dat %in% item_order])
      }

      item_order <- match(item_order, dat)

    } else {
      # add remaining items
      if (length(item_order) < length(dat)) {
        vals <- seq_along(dat)
        item_order <- c(item_order, vals[!vals %in% item_order])
      }
    }

  } else {
    # item_order == NULL
    item_order <- order(dat)
  }

  if (!is.null(choose)) {
    # change order
    choose <- eval(parse(text = choose), data.frame(x = dat))[item_order]
    hide <- as_xml_attr(!choose)
    sele <- as_xml_attr(choose)
  } else {
    hide <- NULL
    sele <- rep("1", length(dat))
  }

  if (slicer) {
    vals <- as.character(item_order - 1L)
    item <- sapply(
      seq_along(vals),
      function(val) {
          xml_node_create("i", xml_attributes = c(x = vals[val], s = sele[val]))
      },
      USE.NAMES = FALSE
    )
  } else {
    vals <- c(as.character(item_order - 1L), default)
    item <- sapply(
      seq_along(vals),
      # # TODO this sets the order of the pivot elements
      # c(seq_along(unique(data[[x]])) - 1L, "default"),
      function(val) {
        if (vals[val] == "default")
          xml_node_create("item", xml_attributes = c(t = vals[val]))
        else
          xml_node_create("item", xml_attributes = c(x = vals[val], h = hide[val]))
      },
      USE.NAMES = FALSE
    )
  }

  items <- xml_node_create(
    "items",
    xml_attributes = c(count = as.character(length(item))), xml_children = item
  )

  items
}

row_col_items <- function(data, z, var) {
  var <- abs(var)
  item <- sapply(
    c(seq_along(unique(data[, var])) - 1L, "grand"),
    function(val) {
      if (val == "0") {
        xml_node_create("i", xml_children = xml_node_create("x"))
      } else if (val == "grand") {
        xml_node_create("i", xml_attributes = c(t = val), xml_children = xml_node_create("x"))
      } else {
        xml_node_create("i", xml_children = xml_node_create("x", xml_attributes = c(v = val)))
      }
    },
    USE.NAMES = FALSE
  )

  xml_node_create(z, xml_attributes = c(count = as.character(length(item))), xml_children = item)
}

create_pivot_table <- function(
    x,
    dims,
    filter,
    rows,
    cols,
    data,
    n,
    fun,
    params,
    numfmts
  ) {

  if (missing(filter)) {
    filter_pos <- 0
    use_filter <- FALSE
    # message("no filter")
  } else {
    filter_pos <- match(filter, names(x))
    use_filter <- TRUE
  }

  if (missing(rows)) {
    rows_pos <- 0
    use_rows <- FALSE
    # message("no rows")
  } else {
    rows_pos <- match(rows, names(x))
    use_rows <- TRUE
  }

  if (missing(cols)) {
    cols_pos <- 0
    use_cols <- FALSE
    # message("no cols")
  } else {
    cols_pos <- match(cols, names(x))
    use_cols <- TRUE
  }

  if (missing(data)) {
    data_pos <- 0
    use_data <- FALSE
    # message("no data")
  } else {
    data_pos <- match(data, names(x))
    use_data <- TRUE
    if (length(data_pos) > 1) {
      cols_pos <- c(-1L, cols_pos[cols_pos > 0])
      use_cols <- TRUE
    }
  }

  pivotField <- NULL
  arguments <- c(
    "apply_alignment_formats", "apply_border_formats", "apply_font_formats",
    "apply_number_formats", "apply_pattern_formats", "apply_width_height_formats",
    "apply_width_height_formats", "asterisk_totals", "auto_format_id",
    "chart_format", "col_grand_totals", "col_header_caption", "compact",
    "compact", "choose", "compact_data", "custom_list_sort", "data_caption",
    "data_on_rows", "data_position", "disable_field_list", "downfill", "edit_data",
    "enable_drill", "enable_field_properties", "enable_wizard", "error_caption",
    "field_list_sort_ascending", "field_print_titles", "grand_total_caption",
    "grid_drop_zones", "immersive", "indent", "item_print_titles",
    "mdx_subqueries", "merge_item", "missing_caption", "multiple_field_filters",
    "name", "no_style", "numfmt", "outline", "outline_data", "page_over_then_down",
    "page_style", "page_wrap", "preserve_formatting", "print_drill",
    "published", "row_grand_totals", "row_header_caption", "show_calc_mbrs",
    "show_col_headers", "show_col_stripes", "show_data_as", "show_data_drop_down",
    "show_data_tips", "show_drill", "show_drop_zones", "show_empty_col",
    "show_empty_row", "show_error", "show_headers", "show_items",
    "show_last_column", "show_member_property_tips", "show_missing",
    "show_multiple_label", "show_row_headers", "show_row_stripes",
    "sort_col", "sort_item", "sort_row", "subtotal_hidden_items",
    "table_style", "tag", "use_auto_formatting", "vacated_style",
    "visual_totals",
    "subtotal_top", "default_subtotal"
  )
  params <- standardize_case_names(params, arguments = arguments, return = TRUE)


  compact <- ""
  if (!is.null(params$compact))
    compact <- params$compact

  outline <- ""
  if (!is.null(params$outline))
    outline <- params$outline

  subtotalTop           <- NULL
  SubtotalsOnTopDefault <- NULL
  if (!is.null(params$subtotal_top)) {
    subtotalTop           <- as_xml_attr(params$subtotal_top)
    SubtotalsOnTopDefault <- as_xml_attr(params$subtotal_top)
  }

  has_default             <- TRUE
  defaultSubtotal         <- NULL
  EnabledSubtotalsDefault <- NULL
  if (!is.null(params$default_subtotal)) {
    has_default             <- params$default_subtotal
    defaultSubtotal         <- as_xml_attr(has_default)
    EnabledSubtotalsDefault <- as_xml_attr(has_default)
  }

  for (i in seq_along(x)) {

    dataField <- NULL
    axis <- NULL
    sort <- NULL
    autoSortScope <- NULL
    downfill <- NULL
    is_formula <- inherits(x[[i]], "is_formula")

    if (i %in% data_pos)    dataField <- c(dataField = "1")

    if (i %in% filter_pos)  axis <- c(axis = "axisPage")

    if (i %in% rows_pos) {
      axis <- c(axis = "axisRow")
      sort <- params$sort_row

      if (!is.null(sort) && !is.character(sort)) {
        if (!abs(sort) %in% seq_along(rows_pos))
          warning("invalid sort position found")

        if (!abs(sort) == match(i, rows_pos))
          sort <- NULL
      }

      if (isTRUE(params$downfill)) {
        downfill <- read_xml(
          '<extLst>
            <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}">
              <x14:pivotField fillDownLabels="1"/>
            </ext>
          </extLst>',
          pointer = FALSE)
      }
    }

    if (i %in% cols_pos) {
      axis <- c(axis = "axisCol")
      sort <- params$sort_col

      if (!is.null(sort) && !is.character(sort)) {
        if (!abs(sort) %in% seq_along(cols_pos))
          warning("invalid sort position found")

        if (!abs(sort) == match(i, cols_pos))
          sort <- NULL
      }
    }

    if (!is.null(sort) && !is.character(sort)) {

      autoSortScope <- read_xml(sprintf('
        <autoSortScope>
          <pivotArea dataOnly="0" outline="0" fieldPosition="0">
            <references count="1">
            <reference field="4294967294" count="1" selected="0">
              <x v="%s" />
            </reference>
            </references>
          </pivotArea>
        </autoSortScope>
        ',
        abs(sort) - 1L), pointer = FALSE)

      if (sign(sort) == -1) sort <- "descending"
      else                  sort <- "ascending"
    }

    sort_item <- params$sort_item
    choose    <- params$choose
    multi     <- if (is.null(choose)) NULL else as_xml_attr(TRUE)

    attrs <- c(
      axis, dataField,
      subtotalTop = subtotalTop,
      showAll = "0",
      defaultSubtotal = defaultSubtotal,
      multipleItemSelectionAllowed = multi,
      sortType = sort,
      compact = as_xml_attr(compact),
      outline = as_xml_attr(outline)
    )

    if (is_formula) {
        attrs <- c(
          dataField       = "1",
          dragToRow       = "0",
          dragToCol       = "0",
          dragToPage      = "0",
          showAll         = "0",
          defaultSubtotal = "0"
        )
    }

    tmp <- xml_node_create(
      "pivotField",
      xml_attributes = attrs)

    if (i %in% c(filter_pos, rows_pos, cols_pos)) {
      nms <- names(x)[i]
      sort_itm <- sort_item[[nms]]
      if (!is.null(choose) && !is.na(choose[nms])) {
        choo <- choose[nms]
      } else {
        choo <- NULL
      }
      tmp <- xml_node_create(
        "pivotField",
        xml_attributes = attrs,
        xml_children = paste0(paste0(get_items(x, i, sort_itm, FALSE, choo, has_default), collapse = ""), autoSortScope, downfill))
    }

    pivotField <- c(pivotField, tmp)
  }

  pivotFields <- xml_node_create(
    "pivotFields",
    xml_children = pivotField,
    xml_attributes = c(count = as.character(length(pivotField))))

  cols_field <- c(cols_pos - 1L)
  # if (length(data_pos) > 1)
  #   cols_field <- c(cols_field * -1, rep(1, length(data_pos) -1L))

  dim <- dims

  if (use_filter) {
    pageFields <- paste0(
      sprintf('<pageFields count="%s">', length(data_pos)),
      paste0(sprintf('<pageField fld="%s" hier="-1" />', filter_pos - 1L), collapse = ""),
      '</pageFields>')
  } else {
    pageFields <- ""
  }

  if (use_cols) {
    colsFields <- paste0(
      sprintf('<colFields count="%s">', length(cols_pos)),
      paste0(sprintf('<field x="%s" />', cols_field), collapse = ""),
      '</colFields>',
      paste0(row_col_items(data = x, "colItems", cols_pos), collapse = "")
    )
  } else {
    colsFields <- ""
  }

  if (use_rows) {
    rowsFields <- paste0(
      sprintf('<rowFields count="%s">', length(rows_pos)),
      paste0(sprintf('<field x="%s" />', rows_pos - 1L), collapse = ""),
      '</rowFields>',
      paste0(row_col_items(data = x, "rowItems", rows_pos), collapse = "")
    )
  } else {
    rowsFields <- ""
  }

  if (use_data) {

    dataField <- NULL

    show_data_as <- rep("", length(data))
    if (!is.null(params$show_data_as))
      show_data_as <- params$show_data_as

    for (i in seq_along(data)) {

      if (missing(fun)) fun <- NULL

      dataField <- c(
        dataField,
        xml_node_create(
          "dataField",
          xml_attributes = c(
            name       = sprintf("%s of %s", ifelse(is.null(fun[i]), "Sum", fun[i]), data[i]),
            fld        = sprintf("%s", data_pos[i] - 1L),
            showDataAs = as_xml_attr(ifelse(is.null(show_data_as[i]), "", show_data_as[i])),
            subtotal   = fun[i],
            baseField  = "0",
            baseItem   = "0",
            numFmtId   = numfmts[i]
          )
        )
      )
    }

    dataFields <- xml_node_create(
      "dataFields",
      xml_attributes = c(count = as.character(length(data_pos))),
      xml_children = paste0(dataField, collapse = "")
    )
  } else {
    dataFields <- ""
  }

  location <- xml_node_create(
    "location",
    xml_attributes = c(
      ref            = dim,
      firstHeaderRow = "1",
      firstDataRow   = "2",
      firstDataCol   = "1",
      rowPageCount   = as_binary(use_filter),
      colPageCount   = as_binary(use_filter)
    )
  )

  table_style <- "PivotStyleLight16"
  if (!is.null(params$table_style))
    table_style <- params$table_style

  dataCaption <- "Values"
  if (!is.null(params$data_caption))
    dataCaption <- params$data_caption

  showRowHeaders <- "1"
  if (!is.null(params$show_row_headers))
    showRowHeaders <- params$show_row_headers

  showColHeaders <- "1"
  if (!is.null(params$show_col_headers))
    showColHeaders <- params$show_col_headers

  showRowStripes <- "0"
  if (!is.null(params$show_row_stripes))
    showRowStripes <- params$show_row_stripes

  showColStripes <- "0"
  if (!is.null(params$show_col_stripes))
    showColStripes <- params$show_col_stripes

  showLastColumn <- "1"
  if (!is.null(params$show_last_column))
    showLastColumn <- params$show_last_column

  pivotTableStyleInfo <- xml_node_create(
    "pivotTableStyleInfo",
    xml_attributes = c(
      name           = table_style,
      showRowHeaders = showRowHeaders,
      showColHeaders = showColHeaders,
      showRowStripes = showRowStripes,
      showColStripes = showColStripes,
      showLastColumn = showLastColumn
    )
  )

  if (isTRUE(params$no_style))
    pivotTableStyleInfo <- ""

  indent <- "0"
  if (!is.null(params$indent))
    indent <- params$indent

  itemPrintTitles <- "1"
  if (!is.null(params$item_print_titles))
    itemPrintTitles <- params$item_print_titles

  multipleFieldFilters <- "0"
  if (!is.null(params$multiple_field_filters))
    multipleFieldFilters <- params$multiple_field_filters

  outlineData <- "1"
  if (!is.null(params$outline_data))
    outlineData <- params$outline_data

  useAutoFormatting <- "1"
  if (!is.null(params$use_auto_formatting))
    useAutoFormatting <- params$use_auto_formatting

  applyAlignmentFormats <- "0"
  if (!is.null(params$apply_alignment_formats))
    applyAlignmentFormats <- params$apply_alignment_formats

  applyNumberFormats <- "0"
  if (!is.null(params$apply_number_formats))
    applyNumberFormats <- params$apply_number_formats

  applyBorderFormats <- "0"
  if (!is.null(params$apply_border_formats))
    applyBorderFormats <- params$apply_border_formats

  applyFontFormats <- "0"
  if (!is.null(params$apply_font_formats))
    applyFontFormats <- params$apply_font_formats

  applyPatternFormats <- "0"
  if (!is.null(params$apply_pattern_formats))
    applyPatternFormats <- params$apply_pattern_formats

  applyWidthHeightFormats <- "1"
  if (!is.null(params$apply_width_height_formats))
    applyWidthHeightFormats <- params$apply_width_height_formats

  pivot_table_name <- sprintf("PivotTable%s", n)
  if (!is.null(params$name))
    pivot_table_name <- params$name


  ptd16 <- xml_node_create(
    "xpdl:pivotTableDefinition16",
    xml_attributes = c(
      EnabledSubtotalsDefault = EnabledSubtotalsDefault,
      SubtotalsOnTopDefault   = SubtotalsOnTopDefault
    )
  )

  extLst <- sprintf(
    '<extLst>
    <ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
    <x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" />
    </ext>
    <ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout">
    %s
    </ext>
    </extLst>',
    ptd16
  )

  xml_node_create(
    "pivotTableDefinition",
    xml_attributes = c(
      xmlns                   = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      `xmlns:mc`              = "http://schemas.openxmlformats.org/markup-compatibility/2006",
      `mc:Ignorable`          = "xr",
      `xmlns:xr`              = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
      name                    = as_xml_attr(pivot_table_name),
      cacheId                 = as_xml_attr(n),
      applyNumberFormats      = as_xml_attr(applyNumberFormats),
      applyBorderFormats      = as_xml_attr(applyBorderFormats),
      applyFontFormats        = as_xml_attr(applyFontFormats),
      applyPatternFormats     = as_xml_attr(applyPatternFormats),
      applyAlignmentFormats   = as_xml_attr(applyAlignmentFormats),
      applyWidthHeightFormats = as_xml_attr(applyWidthHeightFormats),
      asteriskTotals          = as_xml_attr(params$asterisk_totals),
      autoFormatId            = as_xml_attr(params$auto_format_id),
      chartFormat             = as_xml_attr(params$chart_format),
      dataCaption             = as_xml_attr(dataCaption),
      updatedVersion          = "8",
      minRefreshableVersion   = "3",
      useAutoFormatting       = as_xml_attr(useAutoFormatting),
      itemPrintTitles         = as_xml_attr(itemPrintTitles),
      createdVersion          = "8",
      indent                  = as_xml_attr(indent),
      outline                 = as_xml_attr(outline),
      outlineData             = as_xml_attr(outlineData),
      multipleFieldFilters    = as_xml_attr(multipleFieldFilters),
      colGrandTotals          = as_xml_attr(params$col_grand_totals),
      colHeaderCaption        = as_xml_attr(params$col_header_caption),
      compact                 = as_xml_attr(params$compact),
      compactData             = as_xml_attr(params$compact_data),
      customListSort          = as_xml_attr(params$custom_list_sort),
      dataOnRows              = as_xml_attr(params$data_on_rows),
      dataPosition            = as_xml_attr(params$data_position),
      disableFieldList        = as_xml_attr(params$disable_field_list),
      editData                = as_xml_attr(params$edit_data),
      enableDrill             = as_xml_attr(params$enable_drill),
      enableFieldProperties   = as_xml_attr(params$enable_field_properties),
      enableWizard            = as_xml_attr(params$enable_wizard),
      errorCaption            = as_xml_attr(params$error_caption),
      fieldListSortAscending  = as_xml_attr(params$field_list_sort_ascending),
      fieldPrintTitles        = as_xml_attr(params$field_print_titles),
      grandTotalCaption       = as_xml_attr(params$grand_total_caption),
      gridDropZones           = as_xml_attr(params$grid_drop_zones),
      immersive               = as_xml_attr(params$immersive),
      mdxSubqueries           = as_xml_attr(params$mdx_subqueries),
      missingCaption          = as_xml_attr(params$missing_caption),
      mergeItem               = as_xml_attr(params$merge_item),
      pageOverThenDown        = as_xml_attr(params$page_over_then_down),
      pageStyle               = as_xml_attr(params$page_style),
      pageWrap                = as_xml_attr(params$page_wrap),
      # pivotTableStyle
      preserveFormatting      = as_xml_attr(params$preserve_formatting),
      printDrill              = as_xml_attr(params$print_drill),
      published               = as_xml_attr(params$published),
      rowGrandTotals          = as_xml_attr(params$row_grand_totals),
      rowHeaderCaption        = as_xml_attr(params$row_header_caption),
      showCalcMbrs            = as_xml_attr(params$show_calc_mbrs),
      showDataDropDown        = as_xml_attr(params$show_data_drop_down),
      showDataTips            = as_xml_attr(params$show_data_tips),
      showDrill               = as_xml_attr(params$show_drill),
      showDropZones           = as_xml_attr(params$show_drop_zones),
      showEmptyCol            = as_xml_attr(params$show_empty_col),
      showEmptyRow            = as_xml_attr(params$show_empty_row),
      showError               = as_xml_attr(params$show_error),
      showHeaders             = as_xml_attr(params$show_headers),
      showItems               = as_xml_attr(params$show_items),
      showMemberPropertyTips  = as_xml_attr(params$show_member_property_tips),
      showMissing             = as_xml_attr(params$show_missing),
      showMultipleLabel       = as_xml_attr(params$show_multiple_label),
      subtotalHiddenItems     = as_xml_attr(params$subtotal_hidden_items),
      tag                     = as_xml_attr(params$tag),
      vacatedStyle            = as_xml_attr(params$vacated_style),
      visualTotals            = as_xml_attr(params$visual_totals)
    ),
    # xr:uid="{375073AB-E7CA-C149-922E-A999C47476C1}"
    xml_children =  paste0(
      location,
      paste0(pivotFields, collapse = ""),
      rowsFields,
      colsFields,
      pageFields,
      dataFields,
      pivotTableStyleInfo,
      extLst
    )
  )

}
