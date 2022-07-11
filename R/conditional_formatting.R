#' conditional formatting rules
#' @name cf_rules
#' @param formula formula
#' @param values values
#' @keywords internal
#' @noRd
cf_create_colorscale <- function(formula, values) {

  ## formula contains the colours
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
#' @param extLst extLst
#' @param params params
#' @param sqref sqref
#' @keywords internal
#' @details `cf_create_databar()` returns extLst for worksheet
#' @noRd
cf_create_databar <- function(extLst, formula, params, sqref, values) {
  if (length(formula) == 2L) {
    negColour <- formula[[1]]
    posColour <- formula[[2]]
  } else {
    posColour <- formula
    negColour <- "FFFF0000"
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
    posColour = posColour,
    negColour = negColour,
    values    = values,
    border    = border,
    gradient  = gradient
  )

  # check if any extLst availaible
  if (length(extLst) == 0) {
    extLst <- newExtLst
  } else if (length(xml_node(extLst, "ext", "x14:conditionalFormattings")) == 0) {
    # extLst is available, has no conditionalFormattings
    extLst <- xml_add_child(
      extLst,
      xml_node(newExtLst, "ext", "x14:conditionalFormattings")
    )
  } else {
    # extLst is available, has conditionalFormattings
    extLst <- xml_add_child(
      extLst,
      xml_node(
        newExtLst,
        "ext",
        "x14:conditionalFormattings",
        "x14:conditionalFormatting"
      ),
      level = "x14:conditionalFormattings"
    )
  }

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
      posColour,
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
      posColour,
      # extLst
      cf_rule_extLst
    )
  }

  attr(cf_rule, "extLst") <- extLst
  cf_rule
}

#' @rdname cf_rules
#' @param dxfId dxfId
#' @param formula formula
#' @keywords internal
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
#' @keywords internal
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
#' @keywords internal
#' @noRd
cf_create_contains_text <- function(dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="containsText" dxfId="%s" priority="1" operator="containsText" text="%s">
      <formula>NOT(ISERROR(SEARCH("%s", %s)))</formula>
    </cfRule>',
    # cfRule
    dxfId,
    values,
    # formula
    values,
    strsplit(sqref, split = ":")[[1]][1]
  )

  cf_rule
}

#' @rdname cf_rules
#' @keywords internal
#' @noRd
cf_create_not_contains_text <- function(dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="notContainsText" dxfId="%s" priority="1" operator="notContains" text="%s">
      <formula>ISERROR(SEARCH("%s", %s))</formula>
    </cfRule>',
    # cfRule
    dxfId,
    values,
    # formula
    values,
    strsplit(sqref, split = ":")[[1]][1]
  )

  cf_rule
}

#' @rdname cf_rules
#' @keywords internal
#' @noRd
cf_begins_with <- function(dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="beginsWith" dxfId="%s" priority="1" operator="beginsWith" text="%s">
      <formula>LEFT(%s,LEN("%s"))="%s"</formula>
    </cfRule>',
    # cfRule
    dxfId,
    values,
    # formula
    strsplit(sqref, split = ":")[[1]][1],
    values,
    values
  )

  cf_rule
}

#' @rdname cf_rules
#' @keywords internal
#' @noRd
cf_ends_with <- function(dxfId, sqref, values) {
  cf_rule <- sprintf(
    '<cfRule type="endsWith" dxfId="%s" priority="1" operator="endsWith" text="%s">
      <formula>RIGHT(%s,LEN("%s"))="%s"</formula>
    </cfRule>',
    # cfRule
    dxfId,
    values,
    # formula
    strsplit(sqref, split = ":")[[1]][1],
    values,
    values
  )

  cf_rule
}

#' @rdname cf_rules
#' @keywords internal
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
#' @keywords internal
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
#' @keywords internal
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
