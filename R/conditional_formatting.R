#' turn a possibly non consecutive `dims` into the ranges of a single `sqref`
#'
#' The selection is reduced to its set of cells and rebuilt into ranges
#' deterministically, so the result depends only on which cells are selected,
#' not on how `dims` was written. That gives two guarantees:
#' - no cell is named twice, which matters because Excel counts a repeated
#'   cell twice (a `duplicatedValues` rule would mark it as its own duplicate)
#' - the same selection always yields the same `sqref`, which is what
#'   `remove_conditional_formatting()` relies on to find a stored rule
#'
#' @param dims a dims string with one or several blocks, separated by `,`, `;`
#'   or the space of a stored `sqref`
#' @returns `NULL` if `dims` selects no cell, otherwise a list of the `sqref`
#'   and the top left cell of its first range, the `anchor`
#' @noRd
cf_dims_to_sqref <- function(dims) {

  pieces <- unlist(strsplit(dims, split = "[,; ]"))
  pieces <- pieces[nzchar(pieces)]

  if (!length(pieces)) return(NULL)

  # a single block is already a rectangle, avoid expanding it to cells
  if (length(pieces) == 1L) {
    d <- dims_to_rowcol(pieces, as_integer = FALSE)
    rows <- range(as.integer(d$row))
    cols <- range(col2int(d$col))
    return(list(
      sqref  = rowcol_to_dims(rows, cols),
      anchor = rowcol_to_dim(rows[1], cols[1])
    ))
  }

  # one row per selected cell; unique() removes overlaps
  cells <- lapply(pieces, function(x) {
    d <- dims_to_rowcol(x, as_integer = FALSE)
    expand.grid(row = as.integer(d$row), col = col2int(d$col))
  })
  cells <- unique(do.call(rbind, cells))
  cells <- cells[order(cells$col, cells$row), ]

  # split each column into runs of consecutive rows
  first <- which(c(TRUE, diff(cells$col) != 0 | diff(cells$row) != 1))
  last <- c(first[-1] - 1L, nrow(cells))
  runs <- data.frame(r1 = cells$row[first], r2 = cells$row[last], c1 = cells$col[first])

  # neighboring columns with identical runs are joined into one rectangle
  runs <- runs[order(runs$r1, runs$r2, runs$c1), ]
  grp <- cumsum(c(TRUE, diff(runs$r1) != 0 | diff(runs$r2) != 0 | diff(runs$c1) != 1))
  rng <- do.call(rbind, lapply(split(runs, grp), function(g) {
    data.frame(r1 = g$r1[1], r2 = g$r2[1], c1 = g$c1[1], c2 = g$c1[nrow(g)])
  }))

  # top to bottom, left to right: relative references in a rule are anchored
  # to the top left cell of the first range
  rng <- rng[order(rng$r1, rng$c1), ]

  list(
    sqref  = paste(mapply(function(r1, r2, c1, c2) rowcol_to_dims(c(r1, r2), c(c1, c2)),
                          rng$r1, rng$r2, rng$c1, rng$c2), collapse = " "),
    anchor = rowcol_to_dim(rng$r1[1], rng$c1[1])
  )
}

#' conditional formatting rules
#' @name cf_rules
#' @param formula formula
#' @param values values
#' @noRd
cf_create_colorscale <- function(priority, formula, values) {

  ## formula contains the colors
  ## values contains numerics or is NULL

  if (is.null(values)) {
    # could use a switch() here for length to also check against other
    # lengths, if these aren't checked somewhere already?
    if (length(formula) == 2L) {
      cf_rule <- sprintf(
        '<cfRule type="colorScale" priority="%s">
          <colorScale>
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
          </colorScale>
        </cfRule>',
        priority,
        formula[[1]],
        formula[[2]]
      )
    } else if (length(formula) == 3L) {
      cf_rule <- sprintf(
        '<cfRule type="colorScale" priority="%s">
          <colorScale>
            <cfvo type="min"/>
            <cfvo type="percentile" val="50"/>
            <cfvo type="max"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
          </colorScale>
        </cfRule>',
        priority,
        formula[[1]],
        formula[[2]],
        formula[[3]]
      )
    }
  } else {
    if (length(formula) == 2L && length(values) == 2L) {
      cf_rule <- sprintf(
        '<cfRule type="colorScale" priority="%s">
          <colorScale>
            <cfvo type="num" val="%s"/>
            <cfvo type="num" val="%s"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
          </colorScale>
        </cfRule>',
        priority,
        values[[1]],
        values[[2]],
        formula[[1]],
        formula[[2]]
      )
    } else if (length(formula) == 3L && length(values) == 3L) {
      cf_rule <- sprintf(
        '<cfRule type="colorScale" priority="%s">
          <colorScale>
            <cfvo type="num" val="%s"/>
            <cfvo type="num" val="%s"/>
            <cfvo type="num" val="%s"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
          </colorScale>
        </cfRule>',
        priority,
        values[[1]],
        values[[2]],
        values[[3]],
        formula[[1]],
        formula[[2]],
        formula[[3]]
      )
    }
  }

  cf_rule
}

#' @rdname cf_rules
#' @details `cf_create_databar()` returns extLst for worksheet
#' @param extLst extLst
#' @param params params
#' @param sqref sqref
#' @noRd
cf_create_databar <- function(priority, extLst, formula, params, sqref, values) {

  # TODO why is priority passed to this function?
  if (length(formula) == 2L) {
    negColor <- formula[[1]]
    posColor <- formula[[2]]
  } else {
    posColor <- formula
    negColor <- "FFFF0000"
  }

  guid <- stringi::stri_join(
    "F7189283-14F7-4DE0-9601-54DE9DB",
    40000L + length(xml_node(
      extLst,
      "ext",
      "x14:conditionalFormattings",
      "x14:conditionalFormatting"
    ))
  )

  showValue <- as.integer(params$showValue %||% 1L)

  newExtLst <- gen_databar_extlst(
    guid = guid,
    sqref = sqref,
    posColor = posColor,
    negColor = negColor,
    values = values,
    params = params
  )

  cf_rule_extLst <- sprintf(
    '<extLst>
      <ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
        <x14:id>{%s}</x14:id>
      </ext>
    </extLst>',
    guid
  )

  if (is.null(values)) {
    cf_rule <- sprintf(
      '<cfRule type="dataBar" priority="%s">
        <dataBar showValue="%s">
          <cfvo type="min"/>
          <cfvo type="max"/>
          <color rgb="%s"/>
        </dataBar>
        %s
      </cfRule>',
      # dataBar
      priority,
      showValue,
      # color
      posColor,
      # extLst
      cf_rule_extLst
    )
  } else {
    cf_rule <- sprintf(
      '<cfRule type="dataBar" priority="%s">
        <dataBar showValue="%s">
          <cfvo type="num" val="%s"/>
          <cfvo type="num" val="%s"/>
          <color rgb="%s"/>
        </dataBar>
        %s
      </cfRule>',
      # dataBar
      priority,
      showValue,
      # cfvo
      values[[1]],
      values[[2]],
      # color
      posColor,
      # extLst
      cf_rule_extLst
    )
  }

  attr(cf_rule, "extLst") <- newExtLst
  cf_rule
}

#' @rdname cf_rules
#' @param dxfId dxfId
#' @param formula formula
#' @noRd
cf_create_expression <- function(priority, dxfId, formula) {
  cf_rule <- sprintf(
    '<cfRule type="expression" dxfId="%s" priority="%s">
      <formula>%s</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    # formula
    formula
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_create_duplicated_values <- function(priority, dxfId) {
  cf_rule <- sprintf(
    '<cfRule type="duplicateValues" dxfId="%s" priority="%s"/>',
    # cfRule
    dxfId,
    priority
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_create_contains_text <- function(priority, dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="containsText" dxfId="%s" priority="%s" operator="containsText" text="%s">
      <formula>NOT(ISERROR(SEARCH("%s", %s)))</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    replace_legal_chars(values),
    # formula
    replace_legal_chars(values),
    strsplit(sqref, split = ":")[[1]][1]
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_create_not_contains_text <- function(priority, dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="notContainsText" dxfId="%s" priority="%s" operator="notContains" text="%s">
      <formula>ISERROR(SEARCH("%s", %s))</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    replace_legal_chars(values),
    # formula
    replace_legal_chars(values),
    strsplit(sqref, split = ":")[[1]][1]
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_begins_with <- function(priority, dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="beginsWith" dxfId="%s" priority="%s" operator="beginsWith" text="%s">
      <formula>LEFT(%s,LEN("%s"))="%s"</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    replace_legal_chars(values),
    # formula
    strsplit(sqref, split = ":")[[1]][1],
    replace_legal_chars(values),
    replace_legal_chars(values)
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_ends_with <- function(priority, dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="endsWith" dxfId="%s" priority="%s" operator="endsWith" text="%s">
      <formula>RIGHT(%s,LEN("%s"))="%s"</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    replace_legal_chars(values),
    # formula
    strsplit(sqref, split = ":")[[1]][1],
    replace_legal_chars(values),
    replace_legal_chars(values)
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_between <- function(priority, dxfId, formula) {
  cf_rule <- sprintf(
    '<cfRule type="cellIs" dxfId="%s" priority="%s" operator="between">
      <formula>%s</formula>
      <formula>%s</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    # formula
    formula[1],
    formula[2]
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_top_n <- function(priority, dxfId, values) {
  cf_rule <- sprintf(
    '<cfRule type="top10" dxfId="%s" priority="%s" rank="%s" percent="%s"/>',
    # cfRule
    dxfId,
    priority,
    values$rank,
    values$percent
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_bottom_n <- function(priority, dxfId, values) {
  cf_rule <- sprintf(
    '<cfRule type="top10" dxfId="%s" priority="%s" rank="%s" percent="%s" bottom="1"/>',
    # cfRule
    dxfId,
    priority,
    values$rank,
    values$percent
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_icon_set <- function(
    priority,
    extLst,
    sqref,
    values,
    params
  ) {

  type      <- ifelse(params$percent, "percent", "num")
  showValue <- NULL
  reverse   <- NULL
  iconSet   <- NULL

  # per default iconSet creation is store in $conditionalFormatting.
  # The few exceptions are stored in extLst
  guid <- NULL
  x14_ns <- NULL
  if (any(params$iconSet %in% c("3Stars", "3Triangles", "5Boxes", "NoIcons"))) {
    guid <- st_guid()
    x14_ns <- "x14:"
  }

  if (!is.null(params$iconSet))
    iconSet <- params$iconSet

  # only if non default
  if (!is.null(params$showValue))
    if (!params$showValue) showValue <- "0"

  if (!is.null(params$reverse))
    if (params$reverse) reverse <- "1"

  # create cfRule with iconset and cfvo

  cf_rule <- xml_node_create(
    paste0(x14_ns, "cfRule"),
    xml_attributes = c(
      type     = "iconSet",
      priority = as_xml_attr(priority),
      id = guid
    )
  )

  iconset <- xml_node_create(
    paste0(x14_ns, "iconSet"),
    xml_attributes = c(
      iconSet   = iconSet,
      showValue = showValue,
      reverse   = reverse
    )
  )

  for (i in seq_along(values)) {
    if (is.null(x14_ns)) {
      iconset <- xml_add_child(
        iconset,
        xml_child = c(
          xml_node_create(
            "cfvo",
            xml_attributes = c(
              type = type,
              val = values[i]
            )
          )
        )
      )
    } else {
      iconset <- xml_add_child(
        iconset,
        xml_child = c(
          xml_node_create(
            "x14:cfvo",
            xml_attributes = c(
              type = type
            ),
            xml_children = xml_node_create("xm:f",
              xml_children = values[i]
            )
          )
        )
      )
    }
  }

  # return
  xml <- xml_add_child(
    cf_rule,
    xml_child = iconset
  )

  if (!is.null(x14_ns)) {
    extLst <- paste0(
      "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">",
      xml,
      "<xm:sqref>",
      sqref,
      "</xm:sqref>",
      "</x14:conditionalFormatting>"
    )

    xml <- character()
    attr(xml, "extLst") <- extLst

  }

  xml
}

#' @rdname cf_rules
#' @noRd
cf_unique_values <- function(priority, dxfId) {
  cf_rule <- sprintf(
    '<cfRule type="uniqueValues" dxfId="%s" priority="%s"/>',
    dxfId,
    priority
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_iserror <- function(priority, dxfId, sqref) {
  cf_rule <- sprintf(
    '<cfRule type="containsErrors" dxfId="%s" priority="%s">
      <formula>ISERROR(%s)</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    # formula
    sqref
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_isnoerror <- function(priority, dxfId, sqref) {
  cf_rule <- sprintf(
    '<cfRule type="notContainsErrors" dxfId="%s" priority="%s">
      <formula>NOT(ISERROR(%s))</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    # formula
    sqref
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_isblank <- function(priority, dxfId, sqref) {
  cf_rule <- sprintf(
    '<cfRule type="containsBlanks" dxfId="%s" priority="%s">
      <formula>LEN(TRIM(%s))=0</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    # formula
    sqref
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_isnoblank <- function(priority, dxfId, sqref) {
  cf_rule <- sprintf(
    '<cfRule type="notContainsBlanks" dxfId="%s" priority="%s">
      <formula>LEN(TRIM(%s))>0</formula>
    </cfRule>',
    # cfRule
    dxfId,
    priority,
    # formula
    sqref
  )

  cf_rule
}
