wb_protect_impl <- function(
    self,
    private,
    protect             = TRUE,
    password            = NULL,
    lockStructure       = FALSE,
    lockWindows         = FALSE,
    type                = c("1", "2", "4", "8"),
    fileSharing         = FALSE,
    username            = unname(Sys.info()["user"]),
    readOnlyRecommended = FALSE
) {
  if (!protect) {
    self$workbook$workbookProtection <- NULL
    return(self)
  }

  # match.arg() doesn't handle numbers too well
  type <- if (!is.character(type)) as.character(type)
  password <- if (is.null(password)) "" else hashPassword(password)

  # TODO: Shall we parse the existing protection settings and preserve all
  # unchanged attributes?

  if (fileSharing) {
    self$workbook$fileSharing <- xml_node_create(
      "fileSharing",
      xml_attributes = c(
        userName = username,
        readOnlyRecommended = if (readOnlyRecommended | type == "2") "1",
        reservationPassword = password
      )
    )
  }

  self$workbook$workbookProtection <- xml_node_create(
    "workbookProtection",
    xml_attributes = c(
      hashPassword = password,
      lockStructure = toString(as.numeric(lockStructure)),
      lockWindows = toString(as.numeric(lockWindows))
    )
  )

  self$workbook$apps <- xml_node_create("DocSecurity", type)
  invisible(self)
}

wb_protect_worksheet_impl <- function(
    self,
    private,
    sheet      = current_sheet(),
    protect    = TRUE,
    password   = NULL,
    properties = NULL
) {
  sheet <- wb_validate_sheet(self, sheet)

  if (!protect) {
    # initializes as character()
    self$worksheets[[sheet]]$sheetProtection <- character()
    return(self)
  }

  all_props <- worksheet_lock_properties()

  if (!is.null(properties)) {
    # ensure only valid properties are listed
    properties <- match.arg(properties, all_props, several.ok = TRUE)
  }

  properties <- as.character(as.numeric(all_props %in% properties))
  names(properties) <- all_props

  if (!is.null(password))
    properties <- c(properties, password = hashPassword(password))

  self$worksheets[[sheet]]$sheetProtection <- xml_node_create(
    "sheetProtection",
    xml_attributes = c(
      sheet = "1",
      properties[properties != "0"]
    )
  )

  invisible(self)
}
