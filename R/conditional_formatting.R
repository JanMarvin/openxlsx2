#' conditional formatting rules
#' @name cf_rules
#' @param formula formula
#' @param values values
#' @noRd
cf_create_colorscale <- function(formula, values) {

  ## formula contains the colors
  ## values contains numerics or is NULL

  if (is.null(values)) {
    # could use a switch() here for length to also check against other
    # lengths, if these aren't checked somewhere already?
    if (length(formula) == 2L) {
      cf_rule <- sprintf(
        '<cfRule type="colorScale" priority="1">
          <colorScale>
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
          </colorScale>
        </cfRule>',
        formula[[1]],
        formula[[2]]
      )
    } else if (length(formula) == 3L) {
      cf_rule <- sprintf(
        '<cfRule type="colorScale" priority="1">
          <colorScale>
            <cfvo type="min"/>
            <cfvo type="percentile" val="50"/>
            <cfvo type="max"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
          </colorScale>
        </cfRule>',
        formula[[1]],
        formula[[2]],
        formula[[3]]
      )
    }
  } else {
    if (length(formula) == 2L && length(values) == 2L) {
      cf_rule <- sprintf(
        '<cfRule type="colorScale" priority="1">
          <colorScale>
            <cfvo type="num" val="%s"/>
            <cfvo type="num" val="%s"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
          </colorScale>
        </cfRule>',
        values[[1]],
        values[[2]],
        formula[[1]],
        formula[[2]]
      )
    } else if (length(formula) == 3L && length(values) == 3L) {
      cf_rule <- sprintf(
        '<cfRule type="colorScale" priority="1">
          <colorScale>
            <cfvo type="num" val="%s"/>
            <cfvo type="num" val="%s"/>
            <cfvo type="num" val="%s"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
            <color rgb="%s"/>
          </colorScale>
        </cfRule>',
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
cf_create_databar <- function(extLst, formula, params, sqref, values) {
  if (length(formula) == 2L) {
    negColor <- formula[[1]]
    posColor <- formula[[2]]
  } else {
    posColor <- formula
    negColor <- "FFFF0000"
  }

  guid <- stri_join(
    "F7189283-14F7-4DE0-9601-54DE9DB",
    40000L + length(xml_node(
      extLst,
      "ext",
      "x14:conditionalFormattings",
      "x14:conditionalFormatting"
    ))
  )

  showValue <- as.integer(params$showValue %||% 1L)
  gradient  <- as.integer(params$gradient  %||% 1L)
  border    <- as.integer(params$border    %||% 1L)

  newExtLst <- gen_databar_extlst(
    guid      = guid,
    sqref     = sqref,
    posColor  = posColor,
    negColor  = negColor,
    values    = values,
    border    = border,
    gradient  = gradient
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
      '<cfRule type="dataBar" priority="1">
        <dataBar showValue="%s">
          <cfvo type="min"/>
          <cfvo type="max"/>
          <color rgb="%s"/>
        </dataBar>
        %s
      </cfRule>',
      # dataBar
      showValue,
      # color
      posColor,
      # extLst
      cf_rule_extLst
    )
  } else {
    cf_rule <- sprintf(
      '<cfRule type="dataBar" priority="1">
        <dataBar showValue="%s">
          <cfvo type="num" val="%s"/>
          <cfvo type="num" val="%s"/>
          <color rgb="%s"/>
        </dataBar>
        %s
      </cfRule>',
      # dataBar
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
cf_create_expression <- function(dxfId, formula) {
  cf_rule <- sprintf(
    '<cfRule type="expression" dxfId="%s" priority="1">
      <formula>%s</formula>
    </cfRule>',
    # cfRule
    dxfId,
    # formula
    formula
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_create_duplicated_values <- function(dxfId) {
  cf_rule <- sprintf(
    '<cfRule type="duplicateValues" dxfId="%s" priority="1"/>',
    # cfRule
    dxfId
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_create_contains_text <- function(dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="containsText" dxfId="%s" priority="1" operator="containsText" text="%s">
      <formula>NOT(ISERROR(SEARCH("%s", %s)))</formula>
    </cfRule>',
    # cfRule
    dxfId,
    replace_legal_chars(values),
    # formula
    replace_legal_chars(values),
    strsplit(sqref, split = ":")[[1]][1]
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_create_not_contains_text <- function(dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="notContainsText" dxfId="%s" priority="1" operator="notContains" text="%s">
      <formula>ISERROR(SEARCH("%s", %s))</formula>
    </cfRule>',
    # cfRule
    dxfId,
    replace_legal_chars(values),
    # formula
    replace_legal_chars(values),
    strsplit(sqref, split = ":")[[1]][1]
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_begins_with <- function(dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="beginsWith" dxfId="%s" priority="1" operator="beginsWith" text="%s">
      <formula>LEFT(%s,LEN("%s"))="%s"</formula>
    </cfRule>',
    # cfRule
    dxfId,
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
cf_ends_with <- function(dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="endsWith" dxfId="%s" priority="1" operator="endsWith" text="%s">
      <formula>RIGHT(%s,LEN("%s"))="%s"</formula>
    </cfRule>',
    # cfRule
    dxfId,
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
cf_between <- function(dxfId, formula) {
  cf_rule <- sprintf(
    '<cfRule type="cellIs" dxfId="%s" priority="1" operator="between">
      <formula>%s</formula>
      <formula>%s</formula>
    </cfRule>',
    # cfRule
    dxfId,
    # formula
    formula[1],
    formula[2]
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_top_n <- function(dxfId, values) {
  cf_rule <- sprintf(
    '<cfRule type="top10" dxfId="%s" priority="1" rank="%s" percent="%s"/>',
    # cfRule
    dxfId,
    values$rank,
    values$percent
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_bottom_n <- function(dxfId, values) {
  cf_rule <- sprintf(
    '<cfRule type="top10" dxfId="%s" priority="1" rank="%s" percent="%s" bottom="1"/>',
    # cfRule
    dxfId,
    values$rank,
    values$percent
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_icon_set <- function(
    extLst,
    sqref,
    values,
    params
  ) {

  type      <- ifelse(params$percent, "percent", "num")
  priority  <- "1"
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
      priority = priority,
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
cf_unique_values <- function(dxfId) {
  cf_rule <- sprintf(
    '<cfRule type="uniqueValues" dxfId="%s" priority="1"/>',
    dxfId
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_iserror <- function(dxfId, sqref) {
  cf_rule <- sprintf(
    '<cfRule type="containsErrors" dxfId="%s" priority="1">
      <formula>ISERROR(%s)</formula>
    </cfRule>',
    # cfRule
    dxfId,
    # formula
    sqref
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_isnoerror <- function(dxfId, sqref) {
  cf_rule <- sprintf(
    '<cfRule type="notContainsErrors" dxfId="%s" priority="1">
      <formula>NOT(ISERROR(%s))</formula>
    </cfRule>',
    # cfRule
    dxfId,
    # formula
    sqref
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_isblank <- function(dxfId, sqref) {
  cf_rule <- sprintf(
    '<cfRule type="containsBlanks" dxfId="%s" priority="1">
      <formula>LEN(TRIM(%s))=0</formula>
    </cfRule>',
    # cfRule
    dxfId,
    # formula
    sqref
  )

  cf_rule
}

#' @rdname cf_rules
#' @noRd
cf_isnoblank <- function(dxfId, sqref) {
  cf_rule <- sprintf(
    '<cfRule type="notContainsBlanks" dxfId="%s" priority="1">
      <formula>LEN(TRIM(%s))>0</formula>
    </cfRule>',
    # cfRule
    dxfId,
    # formula
    sqref
  )

  cf_rule
}
