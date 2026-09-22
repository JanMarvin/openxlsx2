#' what is left of a block once an earlier one has had its cells
#'
#' The remainder is cut columns first: the columns the earlier block does not
#' cover keep the full height of this block, and only in the shared columns is
#' what lies above and below it cut off. That is the decomposition a
#' spreadsheet application writes for such a selection.
#'
#' @param b,p blocks, each `list(row = c(first, last), col = <column numbers>)`
#' @returns a list of blocks covering the cells of `b` that are not in `p`,
#'   `list(b)` when the two do not overlap, an empty list when `p` covers `b`
#' @noRd
cf_block_minus <- function(b, p) {

  shared <- intersect(b[["col"]], p[["col"]])

  if (!length(shared) || b[["row"]][1] > p[["row"]][2] || b[["row"]][2] < p[["row"]][1])
    return(list(b))

  out <- list()
  outside <- setdiff(b[["col"]], p[["col"]])

  if (length(outside))
    out[[length(out) + 1L]] <- list(row = b[["row"]], col = outside)

  if (b[["row"]][2] > p[["row"]][2])
    out[[length(out) + 1L]] <- list(row = c(p[["row"]][2] + 1L, b[["row"]][2]), col = shared)

  out
}

#' cut overlapping blocks apart so that no cell is listed twice
#'
#' The standard allows a `sqref` to name a cell more than once (`ST_Sqref` is
#' a plain list of `ST_Ref`), but Excel never writes one that does: for
#' `"A1:C3,B2:D4"` it writes `A1:C3 B4:C4 D2:D4`, the second block minus what
#' the first one already covers. We write the same ranges, ordered top to
#' bottom and left to right rather than in the order they were cut, so
#' `"A1:C3 D2:D4 B4:C4"`. Repeating a cell is not cosmetic, Excel counts it
#' twice: with every value distinct, a `duplicatedValues` rule over
#' `A1:C3 B2:D4` marks the four repeated cells as duplicates of themselves and
#' `uniqueValues` drops them.
#'
#' Blocks are taken top to bottom, left to right, and each keeps only the part
#' no earlier block covers, so the result does not depend on the order the
#' blocks were written in.
#'
#' @param blocks list of blocks, each `list(row = c(first, last), col = <cols>)`
#' @returns the same blocks with every overlap removed, and unchanged when
#'   they are disjoint already
#' @noRd
cf_cut_overlaps <- function(blocks) {

  if (length(blocks) < 2L) return(blocks)

  row_beg <- vapply(blocks, function(b) b[["row"]][1], NA_integer_)
  row_end <- vapply(blocks, function(b) b[["row"]][2], NA_integer_)
  col_beg <- vapply(blocks, function(b) min(b[["col"]]), NA_integer_)
  col_end <- vapply(blocks, function(b) max(b[["col"]]), NA_integer_)

  # A total order, not just the top left corner: two blocks can start in the
  # same cell, and order() would then keep them in the order they were
  # written, which is what this must not depend on.
  ord <- order(row_beg, col_beg, row_end, col_end)

  # Sweep down the sheet. `active` holds the earlier blocks that reach the
  # current row; a block whose last row lies above the current first row is
  # dropped for good, because every later block starts no higher. Disjoint
  # blocks, the usual case, keep this set at a handful of entries.

  kept <- list()
  active <- integer()

  for (i in ord) {
    active <- active[row_end[active] >= row_beg[i]]

    parts <- list(blocks[[i]])

    # Cutting against the earlier blocks themselves rather than against the
    # parts they were reduced to gives the same cells: the parts kept so far
    # cover exactly the union of those blocks.
    for (j in active) {
      if (col_end[j] < col_beg[i] || col_beg[j] > col_end[i]) next

      parts <- unlist(lapply(parts, cf_block_minus, p = blocks[[j]]), recursive = FALSE)
      if (!length(parts)) break
    }

    kept <- c(kept, parts)
    active <- c(active, i)
  }

  kept
}

#' turn a possibly non consecutive `dims` into the ranges of a single `sqref`
#'
#' Blocks are kept apart and merged back together along one axis whenever they
#' agree on the other one, which never changes the selected cells: same columns
#' and touching rows, or same rows and any columns. That keeps `"A2,A3,A5"`
#' collapsing into A2:A3 and A5:A5, and keeps blocks written side by side
#' ("A1:B4,C1:C4") the single range that pooling used to produce. Merging on
#' one axis can enable a merge on the other, so this runs to a fixed point.
#' What is never merged are blocks that differ on both axes: a range the user
#' wrote as one block must not be cut into pieces. Blocks that overlap are the
#' one exception, see [cf_cut_overlaps()].
#'
#' Both adding and removing go through this, so the stored `sqref` is found
#' again whatever spelling the merge can normalise: `"A1:B4,C1:C4"` as well as
#' `"A1:C4"`. A decomposition whose blocks agree on neither axis, a pinwheel
#' tiling of a rectangle, stays several ranges and does not match the single
#' range the same cells were stored as.
#'
#' @param dims a dims string, with or without several blocks
#' @returns `NULL` if `dims` selects no cell at all, otherwise a list of the
#'   `sqref` and the row and column of its first range
#' @noRd
cf_dims_to_sqref <- function(dims) {

  pieces <- unlist(strsplit(dims, split = "[,;]"))
  blocks <- lapply(pieces[nzchar(pieces)], function(x) {
    bdims <- dims_to_rowcol(x, as_integer = FALSE)
    list(
      row = range(as.integer(bdims[["row"]])),
      col = sort(unique(col2int(bdims[["col"]])))
    )
  })

  if (!length(blocks)) return(NULL)

  blocks <- cf_cut_overlaps(blocks)

  repeat {
    n <- length(blocks)

    # same columns, rows touching or overlapping
    key <- vapply(blocks, function(b) paste(b[["col"]], collapse = ","), NA_character_)
    merged <- list()

    for (k in unique(key)) {
      grp <- blocks[key == k]
      grp <- grp[order(vapply(grp, function(b) b[["row"]][1], NA_integer_))]
      cur <- grp[[1]]

      for (i in seq_along(grp)[-1]) {
        nxt <- grp[[i]][["row"]]
        if (nxt[1] <= cur[["row"]][2] + 1L) { # touching or overlapping
          cur[["row"]][2] <- max(cur[["row"]][2], nxt[2])
        } else {
          merged[[length(merged) + 1L]] <- cur
          cur <- grp[[i]]
        }
      }

      merged[[length(merged) + 1L]] <- cur
    }

    # same rows, columns pooled; gaps are split into runs below
    key <- vapply(merged, function(b) paste(b[["row"]], collapse = ":"), NA_character_)
    blocks <- unname(lapply(split(merged, key), function(grp) {
      list(
        row = grp[[1]][["row"]],
        col = sort(unique(unlist(lapply(grp, `[[`, "col"))))
      )
    }))

    if (length(blocks) == n) break
  }

  # top to bottom, left to right, whatever order the blocks were written in.
  # Relative references in a rule are anchored to the top left cell of the
  # first range, as a spreadsheet application does it for such a selection.
  blocks <- blocks[order(
    vapply(blocks, function(b) b[["row"]][1], NA_integer_),
    vapply(blocks, function(b) b[["col"]][1], NA_integer_)
  )]

  # gaps in the columns of a block are split into consecutive runs
  ranges <- unlist(lapply(blocks, function(b) {
    cc <- b[["col"]]
    runs <- unname(tapply(cc, cumsum(c(1, diff(cc) != 1)), function(g) range(g)))
    vapply(runs, function(cr) {
      stringi::stri_join(
        get_cell_refs(data.frame(x = b[["row"]], y = cr, stringsAsFactors = FALSE)),
        collapse = ":"
      )
    }, NA_character_)
  }))

  list(
    sqref = stringi::stri_join(ranges, collapse = " "),
    row   = blocks[[1]][["row"]],
    col   = blocks[[1]][["col"]]
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
