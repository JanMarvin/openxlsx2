workbook_add_border <- function(
    self,
    private,
    sheet         = current_sheet(),
    dims          = "A1",
    bottom_color  = wb_colour(hex = "FF000000"),
    left_color    = wb_colour(hex = "FF000000"),
    right_color   = wb_colour(hex = "FF000000"),
    top_color     = wb_colour(hex = "FF000000"),
    bottom_border = "thin",
    left_border   = "thin",
    right_border  = "thin",
    top_border    = "thin",
    inner_hgrid   = NULL,
    inner_hcolor  = NULL,
    inner_vgrid   = NULL,
    inner_vcolor  = NULL
) {

  # TODO merge styles and if a style is already present, only add the newly
  # created border style

  # cc <- wb$worksheets[[sheet]]$sheet_data$cc
  # df_s <- as.data.frame(lapply(df, function(x) cc$c_s[cc$r %in% x]))

  df <- dims_to_dataframe(dims, fill = TRUE)
  sheet <- private$get_sheet_index(sheet)

  private$do_cell_init(sheet, dims)

  ### beg border creation
  full_single <- create_border(
    top = top_border, top_color = top_color,
    bottom = bottom_border, bottom_color = bottom_color,
    left = left_border, left_color = left_color,
    right = left_border, right_color = right_color
  )

  top_single <- create_border(
    top = top_border, top_color = top_color,
    bottom = inner_hgrid, bottom_color = inner_hcolor,
    left = left_border, left_color = left_color,
    right = left_border, right_color = right_color
  )

  middle_single <- create_border(
    top = inner_hgrid, top_color = inner_hcolor,
    bottom = inner_hgrid, bottom_color = inner_hcolor,
    left = left_border, left_color = left_color,
    right = left_border, right_color = right_color
  )

  bottom_single <- create_border(
    top = inner_hgrid, top_color = inner_hcolor,
    bottom = bottom_border, bottom_color = bottom_color,
    left = left_border, left_color = left_color,
    right = left_border, right_color = right_color
  )

  left_single <- create_border(
    top = top_border, top_color = top_color,
    bottom = bottom_border, bottom_color = bottom_color,
    left = left_border, left_color = left_color,
    right = inner_vgrid, right_color = inner_vcolor
  )

  right_single <- create_border(
    top = top_border, top_color = top_color,
    bottom = bottom_border, bottom_color = bottom_color,
    left = inner_vgrid, left_color = inner_vcolor,
    right = right_border, right_color = right_color
  )

  center_single <- create_border(
    top = top_border, top_color = top_color,
    bottom = bottom_border, bottom_color = bottom_color,
    left = inner_vgrid, left_color = inner_vcolor,
    right = inner_vgrid, right_color = inner_vcolor
  )

  top_left <- create_border(
    top = top_border, top_color = top_color,
    bottom = inner_hgrid, bottom_color = inner_hcolor,
    left = left_border, left_color = left_color,
    right = inner_vgrid, right_color = inner_vcolor
  )

  top_right <- create_border(
    top = top_border, top_color = top_color,
    bottom = inner_hgrid, bottom_color = inner_hcolor,
    left = inner_vgrid, left_color = inner_vcolor,
    right = left_border, right_color = right_color
  )

  bottom_left <- create_border(
    top = inner_hgrid, top_color = inner_hcolor,
    bottom = bottom_border, bottom_color = bottom_color,
    left = left_border, left_color = left_color,
    right = inner_vgrid, right_color = inner_vcolor
  )

  bottom_right <- create_border(
    top = inner_hgrid, top_color = inner_hcolor,
    bottom = bottom_border, bottom_color = bottom_color,
    left = inner_vgrid, left_color = inner_vcolor,
    right = left_border, right_color = right_color
  )

  top_center <- create_border(
    top = top_border, top_color = top_color,
    bottom = inner_hgrid, bottom_color = inner_hcolor,
    left = inner_vgrid, left_color = inner_vcolor,
    right = inner_vgrid, right_color = inner_vcolor
  )

  bottom_center <- create_border(
    top = inner_hgrid, top_color = inner_hcolor,
    bottom = bottom_border, bottom_color = bottom_color,
    left = inner_vgrid, left_color = inner_vcolor,
    right = inner_vgrid, right_color = inner_vcolor
  )

  middle_left <- create_border(
    top = inner_hgrid, top_color = inner_hcolor,
    bottom = inner_hgrid, bottom_color = inner_hcolor,
    left = left_border, left_color = left_color,
    right = inner_vgrid, right_color = inner_vcolor
  )

  middle_right <- create_border(
    top = inner_hgrid, top_color = inner_hcolor,
    bottom = inner_hgrid, bottom_color = inner_hcolor,
    left = inner_vgrid, left_color = inner_vcolor,
    right = right_border, right_color = right_color
  )

  inner_cell <- create_border(
    top = inner_hgrid, top_color = inner_hcolor,
    bottom = inner_hgrid, bottom_color = inner_hcolor,
    left = inner_vgrid, left_color = inner_vcolor,
    right = inner_vgrid, right_color = inner_vcolor
  )
  ### end border creation

  #
  # /* top_single    */
  # /* middle_single */
  # /* bottom_single */
  #

  # /* left_single --- center_single --- right_single */

  #
  # /* top_left   --- top_center   ---  top_right */
  # /*  -                                    -    */
  # /*  -                                    -    */
  # /*  -                                    -    */
  # /* left_middle                   right_middle */
  # /*  -                                    -    */
  # /*  -                                    -    */
  # /*  -                                    -    */
  # /* left_bottom - bottom_center - bottom_right */
  #

  ## beg cell references
  if (ncol(df) == 1 && nrow(df) == 1)
    dim_full_single <- df[1, 1]

  if (ncol(df) == 1 && nrow(df) >= 2) {
    dim_top_single <- df[1, 1]
    dim_bottom_single <- df[nrow(df), 1]
    if (nrow(df) >= 3) {
      mid <- df[, 1]
      dim_middle_single <- mid[!mid %in% c(dim_top_single, dim_bottom_single)]
    }
  }

  if (ncol(df) >= 2 && nrow(df) == 1) {
    dim_left_single <- df[1, 1]
    dim_right_single <- df[1, ncol(df)]
    if (ncol(df) >= 3) {
      ctr <- df[1, ]
      dim_center_single <- ctr[!ctr %in% c(dim_left_single, dim_right_single)]
    }
  }

  if (ncol(df) >= 2 && nrow(df) >= 2) {
    dim_top_left     <- df[1, 1]
    dim_bottom_left  <- df[nrow(df), 1]
    dim_top_right    <- df[1, ncol(df)]
    dim_bottom_right <- df[nrow(df), ncol(df)]

    if (nrow(df) >= 3) {
      top_mid <- df[, 1]
      bottom_mid <- df[, ncol(df)]

      dim_middle_left <- top_mid[!top_mid %in% c(dim_top_left, dim_bottom_left)]
      dim_middle_right <- bottom_mid[!bottom_mid %in% c(dim_top_right, dim_bottom_right)]
    }

    if (ncol(df) >= 3) {
      top_ctr <- df[1, ]
      bottom_ctr <- df[nrow(df), ]

      dim_top_center <- top_ctr[!top_ctr %in% c(dim_top_left, dim_top_right)]
      dim_bottom_center <- bottom_ctr[!bottom_ctr %in% c(dim_bottom_left, dim_bottom_right)]
    }

    if (ncol(df) > 2 && nrow(df) > 2) {
      t_row <- 1
      b_row <- nrow(df)
      l_row <- 1
      r_row <- ncol(df)
      dim_inner_cell <- as.character(unlist(df[c(-t_row, -b_row), c(-l_row, -r_row)]))
    }
  }
  ### end cell references

  # add some random string to the name. if called multiple times, new
  # styles will be created. We do not look for identical styles, therefor
  # we might create duplicates, but if a single style changes, the rest of
  # the workbook remains valid.
  smp <- random_string()
  s <- function(x) paste0(smp, "s", deparse(substitute(x)), seq_along(x))
  sfull_single <- paste0(smp, "full_single")
  stop_single <- paste0(smp, "full_single")
  sbottom_single <- paste0(smp, "bottom_single")
  smiddle_single <- paste0(smp, "middle_single")
  sleft_single <- paste0(smp, "left_single")
  sright_single <- paste0(smp, "right_single")
  scenter_single <- paste0(smp, "center_single")
  stop_left <- paste0(smp, "top_left")
  stop_right <- paste0(smp, "top_right")
  sbottom_left <- paste0(smp, "bottom_left")
  sbottom_right <- paste0(smp, "bottom_right")
  smiddle_left <- paste0(smp, "middle_left")
  smiddle_right <- paste0(smp, "middle_right")
  stop_center <- paste0(smp, "top_center")
  sbottom_center <- paste0(smp, "bottom_center")
  sinner_cell <- paste0(smp, "inner_cell")

  # ncol == 1
  if (ncol(df) == 1) {

    # single cell
    if (nrow(df) == 1) {
      self$styles_mgr$add(full_single, sfull_single)
      xf_prev <- get_cell_styles(self, sheet, dims)
      xf_full_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sfull_single))
      self$styles_mgr$add(xf_full_single, xf_full_single)
      self$set_cell_style(sheet, dims, self$styles_mgr$get_xf_id(xf_full_single))
    }

    # create top & bottom piece
    if (nrow(df) >= 2) {

      # top single
      self$styles_mgr$add(top_single, stop_single)
      xf_prev <- get_cell_styles(self, sheet, dim_top_single)
      xf_top_single <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_single))
      self$styles_mgr$add(xf_top_single, xf_top_single)
      self$set_cell_style(sheet, dim_top_single, self$styles_mgr$get_xf_id(xf_top_single))

      # bottom single
      self$styles_mgr$add(bottom_single, sbottom_single)
      xf_prev <- get_cell_styles(self, sheet, dim_bottom_single)
      xf_bottom_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_single))
      self$styles_mgr$add(xf_bottom_single, xf_bottom_single)
      self$set_cell_style(sheet, dim_bottom_single, self$styles_mgr$get_xf_id(xf_bottom_single))
    }

    # create middle piece(s)
    if (nrow(df) >= 3) {

      # middle single
      self$styles_mgr$add(middle_single, smiddle_single)
      xf_prev <- get_cell_styles(self, sheet, dim_middle_single)
      xf_middle_single <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_single))
      self$styles_mgr$add(xf_middle_single, xf_middle_single)
      self$set_cell_style(sheet, dim_middle_single, self$styles_mgr$get_xf_id(xf_middle_single))
    }

  }

  # create left and right single row pieces
  if (ncol(df) >= 2 && nrow(df) == 1) {

    # left single
    self$styles_mgr$add(left_single, sleft_single)
    xf_prev <- get_cell_styles(self, sheet, dim_left_single)
    xf_left_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sleft_single))
    self$styles_mgr$add(xf_left_single, xf_left_single)
    self$set_cell_style(sheet, dim_left_single, self$styles_mgr$get_xf_id(xf_left_single))

    # right single
    self$styles_mgr$add(right_single, sright_single)
    xf_prev <- get_cell_styles(self, sheet, dim_right_single)
    xf_right_single <- set_border(xf_prev, self$styles_mgr$get_border_id(sright_single))
    self$styles_mgr$add(xf_right_single, xf_right_single)
    self$set_cell_style(sheet, dim_right_single, self$styles_mgr$get_xf_id(xf_right_single))

    # add single center piece(s)
    if (ncol(df) >= 3) {

      # center single
      self$styles_mgr$add(center_single, scenter_single)
      xf_prev <- get_cell_styles(self, sheet, dim_center_single)
      xf_center_single <- set_border(xf_prev, self$styles_mgr$get_border_id(scenter_single))
      self$styles_mgr$add(xf_center_single, xf_center_single)
      self$set_cell_style(sheet, dim_center_single, self$styles_mgr$get_xf_id(xf_center_single))
    }

  }

  # create left & right - top & bottom corners pieces
  if (ncol(df) >= 2 && nrow(df) >= 2) {

    # top left
    self$styles_mgr$add(top_left, stop_left)
    xf_prev <- get_cell_styles(self, sheet, dim_top_left)
    xf_top_left <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_left))
    self$styles_mgr$add(xf_top_left, xf_top_left)
    self$set_cell_style(sheet, dim_top_left, self$styles_mgr$get_xf_id(xf_top_left))

    # top right
    self$styles_mgr$add(top_right, stop_right)
    xf_prev <- get_cell_styles(self, sheet, dim_top_right)
    xf_top_right <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_right))
    self$styles_mgr$add(xf_top_right, xf_top_right)
    self$set_cell_style(sheet, dim_top_right, self$styles_mgr$get_xf_id(xf_top_right))

    # bottom left
    self$styles_mgr$add(bottom_left, sbottom_left)
    xf_prev <- get_cell_styles(self, sheet, dim_bottom_left)
    xf_bottom_left <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_left))
    self$styles_mgr$add(xf_bottom_left, xf_bottom_left)
    self$set_cell_style(sheet, dim_bottom_left, self$styles_mgr$get_xf_id(xf_bottom_left))

    # bottom right
    self$styles_mgr$add(bottom_right, sbottom_right)
    xf_prev <- get_cell_styles(self, sheet, dim_bottom_right)
    xf_bottom_right <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_right))
    self$styles_mgr$add(xf_bottom_right, xf_bottom_right)
    self$set_cell_style(sheet, dim_bottom_right, self$styles_mgr$get_xf_id(xf_bottom_right))
  }

  # create left and right middle pieces
  if (ncol(df) >= 2 && nrow(df) >= 3) {

    # middle left
    self$styles_mgr$add(middle_left, smiddle_left)
    xf_prev <- get_cell_styles(self, sheet, dim_middle_left)
    xf_middle_left <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_left))
    self$styles_mgr$add(xf_middle_left, xf_middle_left)
    self$set_cell_style(sheet, dim_middle_left, self$styles_mgr$get_xf_id(xf_middle_left))

    # middle right
    self$styles_mgr$add(middle_right, smiddle_right)
    xf_prev <- get_cell_styles(self, sheet, dim_middle_right)
    xf_middle_right <- set_border(xf_prev, self$styles_mgr$get_border_id(smiddle_right))
    self$styles_mgr$add(xf_middle_right, xf_middle_right)
    self$set_cell_style(sheet, dim_middle_right, self$styles_mgr$get_xf_id(xf_middle_right))
  }

  # create top and bottom center pieces
  if (ncol(df) >= 3 & nrow(df) >= 2) {

    # top center
    self$styles_mgr$add(top_center, stop_center)
    xf_prev <- get_cell_styles(self, sheet, dim_top_center)
    xf_top_center <- set_border(xf_prev, self$styles_mgr$get_border_id(stop_center))
    self$styles_mgr$add(xf_top_center, xf_top_center)
    self$set_cell_style(sheet, dim_top_center, self$styles_mgr$get_xf_id(xf_top_center))

    # bottom center
    self$styles_mgr$add(bottom_center, sbottom_center)
    xf_prev <- get_cell_styles(self, sheet, dim_bottom_center)
    xf_bottom_center <- set_border(xf_prev, self$styles_mgr$get_border_id(sbottom_center))
    self$styles_mgr$add(xf_bottom_center, xf_bottom_center)
    self$set_cell_style(sheet, dim_bottom_center, self$styles_mgr$get_xf_id(xf_bottom_center))
  }

  if (nrow(df) > 2 && ncol(df) > 2) {

    # inner cells
    self$styles_mgr$add(inner_cell, sinner_cell)
    xf_prev <- get_cell_styles(self, sheet, dim_inner_cell)
    xf_inner_cell <- set_border(xf_prev, self$styles_mgr$get_border_id(sinner_cell))
    self$styles_mgr$add(xf_inner_cell, xf_inner_cell)
    self$set_cell_style(sheet, dim_inner_cell, self$styles_mgr$get_xf_id(xf_inner_cell))
  }

  invisible(self)
}

workbook_add_fill <- function(
    self,
    private,
    sheet         = current_sheet(),
    dims          = "A1",
    color         = wb_colour(hex = "FFFFFF00"),
    pattern       = "solid",
    gradient_fill = "",
    every_nth_col = 1,
    every_nth_row = 1
) {
  sheet <- private$get_sheet_index(sheet)
  private$do_cell_init(sheet, dims)

  # dim in dataframe can contain various styles. go cell by cell.
  did <- dims_to_dataframe(dims, fill = TRUE)
  # select a few cols and rows to fill
  cols <- (seq_len(ncol(did)) %% every_nth_col) == 0
  rows <- (seq_len(nrow(did)) %% every_nth_row) == 0

  dims <- unname(unlist(did[rows, cols, drop = FALSE]))

  cc <- self$worksheets[[sheet]]$sheet_data$cc
  cc <- cc[cc$r %in% dims, ]
  styles <- unique(cc[["c_s"]])

  for (style in styles) {
    dim <- cc[cc$c_s == style, "r"]

    new_fill <- create_fill(
      gradientFill = gradient_fill,
      patternType = pattern,
      fgColor = color
    )
    self$styles_mgr$add(new_fill, new_fill)

    xf_prev <- get_cell_styles(self, sheet, dim[[1]])
    xf_new_fill <- set_fill(xf_prev, self$styles_mgr$get_fill_id(new_fill))
    self$styles_mgr$add(xf_new_fill, xf_new_fill)
    s_id <- self$styles_mgr$get_xf_id(xf_new_fill)
    self$set_cell_style(sheet, dim, s_id)
  }

  invisible(self)
}

workbook_add_font <- function(
    self,
    private,
    sheet     = current_sheet(),
    dims      = "A1",
    name      = "Calibri",
    color     = wb_colour(hex = "FF000000"),
    size      = "11",
    bold      = "",
    italic    = "",
    outline   = "",
    strike    = "",
    underline = "",
    # fine tuning
    charset   = "",
    condense  = "",
    extend    = "",
    family    = "",
    scheme    = "",
    shadow    = "",
    vertAlign = ""
) {
  sheet <- private$get_sheet_index(sheet)
  private$do_cell_init(sheet, dims)

  did <- dims_to_dataframe(dims, fill = TRUE)
  dims <- unname(unlist(did))

  cc <- self$worksheets[[sheet]]$sheet_data$cc
  cc <- cc[cc$r %in% dims, ]
  styles <- unique(cc[["c_s"]])

  for (style in styles) {
    dim <- cc[cc$c_s == style, "r"]

    new_font <- create_font(
      b = bold,
      charset = charset,
      color = color,
      condense = condense,
      extend = extend,
      family = family,
      i = italic,
      name = name,
      outline = outline,
      scheme = scheme,
      shadow = shadow,
      strike = strike,
      sz = size,
      u = underline,
      vertAlign = vertAlign
    )
    self$styles_mgr$add(new_font, new_font)

    xf_prev <- get_cell_styles(self, sheet, dim[[1]])
    xf_new_font <- set_font(xf_prev, self$styles_mgr$get_font_id(new_font))

    self$styles_mgr$add(xf_new_font, xf_new_font)
    s_id <- self$styles_mgr$get_xf_id(xf_new_font)
    self$set_cell_style(sheet, dim, s_id)
  }

  invisible(self)
}

workbook_add_numfmt <- function(
    self,
    private,
    sheet = current_sheet(),
    dims  = "A1",
    numfmt
) {
  sheet <- private$get_sheet_index(sheet)
  private$do_cell_init(sheet, dims)

  did <- dims_to_dataframe(dims, fill = TRUE)
  dims <- unname(unlist(did))

  cc <- self$worksheets[[sheet]]$sheet_data$cc
  cc <- cc[cc$r %in% dims, ]
  styles <- unique(cc[["c_s"]])

  if (inherits(numfmt, "character")) {

    for (style in styles) {
      dim <- cc[cc$c_s == style, "r"]

      new_numfmt <- create_numfmt(
        numFmtId = self$styles_mgr$next_numfmt_id(),
        formatCode = numfmt
      )
      self$styles_mgr$add(new_numfmt, new_numfmt)

      xf_prev <- get_cell_styles(self, sheet, dim[[1]])
      xf_new_numfmt <- set_numfmt(xf_prev, self$styles_mgr$get_numfmt_id(new_numfmt))
      self$styles_mgr$add(xf_new_numfmt, xf_new_numfmt)
      s_id <- self$styles_mgr$get_xf_id(xf_new_numfmt)
      self$set_cell_style(sheet, dim, s_id)
    }

  } else { # format is numeric
    for (style in styles) {
      dim <- cc[cc$c_s == style, "r"]
      xf_prev <- get_cell_styles(self, sheet, dim[[1]])
      xf_new_numfmt <- set_numfmt(xf_prev, numfmt)
      self$styles_mgr$add(xf_new_numfmt, xf_new_numfmt)
      s_id <- self$styles_mgr$get_xf_id(xf_new_numfmt)
      self$set_cell_style(sheet, dim, s_id)
    }

  }

  invisible(self)
}

workbook_add_cell_style <- function(
    self,
    private,
    sheet             = current_sheet(),
    dims              = "A1",
    applyAlignment    = NULL,
    applyBorder       = NULL,
    applyFill         = NULL,
    applyFont         = NULL,
    applyNumberFormat = NULL,
    applyProtection   = NULL,
    borderId          = NULL,
    extLst            = NULL,
    fillId            = NULL,
    fontId            = NULL,
    hidden            = NULL,
    horizontal        = NULL,
    indent            = NULL,
    justifyLastLine   = NULL,
    locked            = NULL,
    numFmtId          = NULL,
    pivotButton       = NULL,
    quotePrefix       = NULL,
    readingOrder      = NULL,
    relativeIndent    = NULL,
    shrinkToFit       = NULL,
    textRotation      = NULL,
    vertical          = NULL,
    wrapText          = NULL,
    xfId              = NULL
) {
  sheet <- private$get_sheet_index(sheet)
  private$do_cell_init(sheet, dims)

  did <- dims_to_dataframe(dims, fill = TRUE)
  dims <- unname(unlist(did))

  cc <- self$worksheets[[sheet]]$sheet_data$cc
  cc <- cc[cc$r %in% dims, ]
  styles <- unique(cc[["c_s"]])

  for (style in styles) {
    dim <- cc[cc$c_s == style, "r"]
    xf_prev <- get_cell_styles(self, sheet, dim[[1]])
    xf_new_cellstyle <- set_cellstyle(
      xf_node           = xf_prev,
      applyAlignment    = applyAlignment,
      applyBorder       = applyBorder,
      applyFill         = applyFill,
      applyFont         = applyFont,
      applyNumberFormat = applyNumberFormat,
      applyProtection   = applyProtection,
      borderId          = borderId,
      extLst            = extLst,
      fillId            = fillId,
      fontId            = fontId,
      hidden            = hidden,
      horizontal        = horizontal,
      indent            = indent,
      justifyLastLine   = justifyLastLine,
      locked            = locked,
      numFmtId          = numFmtId,
      pivotButton       = pivotButton,
      quotePrefix       = quotePrefix,
      readingOrder      = readingOrder,
      relativeIndent    = relativeIndent,
      shrinkToFit       = shrinkToFit,
      textRotation      = textRotation,
      vertical          = vertical,
      wrapText          = wrapText,
      xfId              = xfId
    )
    self$styles_mgr$add(xf_new_cellstyle, xf_new_cellstyle)
    s_id <- self$styles_mgr$get_xf_id(xf_new_cellstyle)
    self$set_cell_style(sheet, dim, s_id)
  }

  invisible(self)
}

workbook_get_cell_style <- function(
    self,
    private,
    sheet = current_sheet(),
    dims
) {
  if (length(dims) == 1 && grepl(":", dims)) {}
    dims <- dims_to_dataframe(dims, fill = TRUE)
  sheet <- private$get_sheet_index(sheet)

  # This alters the workbook
  temp <- self$clone()$.__enclos_env__$private$do_cell_init(sheet, dims)

  # if a range is passed (e.g. "A1:B2") we need to get every cell
  dims <- unname(unlist(dims))

  # TODO check that cc$r is alway valid. not sure atm
  sel <- temp$worksheets[[sheet]]$sheet_data$cc$r %in% dims
  temp$worksheets[[sheet]]$sheet_data$cc$c_s[sel]
}

workbook_set_cell_style <- function(
    self,
    private,
    sheet = current_sheet(),
    dims,
    style
) {
  if (length(dims) == 1 && grepl(":|;", dims))
    dims <- dims_to_dataframe(dims, fill = TRUE)

  sheet <- private$get_sheet_index(sheet)
  private$do_cell_init(sheet, dims)

  # if a range is passed (e.g. "A1:B2") we need to get every cell
  dims <- unname(unlist(dims))
  sel <- self$worksheets[[sheet]]$sheet_data$cc$r %in% dims
  self$worksheets[[sheet]]$sheet_data$cc$c_s[sel] <- style
  invisible(self)
}

workbook_clone_sheet_style = function(
    self,
    private,
    from = current_sheet(),
    to
) {
  id_org <- private$get_sheet_index(from)
  id_new <- private$get_sheet_index(to)

  org_style <- self$worksheets[[id_org]]$sheet_data$cc
  wb_style  <- self$worksheets[[id_new]]$sheet_data$cc

  # only clone styles from sheets with cc
  if (is.null(org_style)) {
    message("'from' has no sheet data styles to clone")
  } else {

    if (is.null(wb_style)) # if null, create empty dataframe
      wb_style <- create_char_dataframe(names(org_style), n = 0)

    # remove all values
    org_style <- org_style[c("r", "row_r", "c_r", "c_s")]

    # do not merge c_s and do not create duplicates
    merged_style <- merge(org_style,
                          wb_style[-which(names(wb_style) == "c_s")],
                          all = TRUE)
    merged_style[is.na(merged_style)] <- ""
    merged_style <- merged_style[!duplicated(merged_style["r"]), ]

    # will be ordere on save
    self$worksheets[[id_new]]$sheet_data$cc <- merged_style

  }

  # copy entire attributes from original sheet to new sheet
  org_rows <- self$worksheets[[id_org]]$sheet_data$row_attr
  new_rows <- self$worksheets[[id_new]]$sheet_data$row_attr

  if (is.null(org_style)) {
    message("'from' has no row styles to clone")
  } else {

    if (is.null(new_rows))
      new_rows <- create_char_dataframe(names(org_rows), n = 0)

    # only add the row information, nothing else
    merged_rows <- merge(org_rows,
                         new_rows["r"],
                         all = TRUE)
    merged_rows[is.na(merged_rows)] <- ""
    merged_rows <- merged_rows[!duplicated(merged_rows["r"]), ]
    ordr <- ordered(order(as.integer(merged_rows$r)))
    merged_rows <- merged_rows[ordr, ]

    self$worksheets[[id_new]]$sheet_data$row_attr <- merged_rows
  }

  self$worksheets[[id_new]]$cols_attr <-
    self$worksheets[[id_org]]$cols_attr

  self$worksheets[[id_new]]$dimension <-
    self$worksheets[[id_org]]$dimension

  self$worksheets[[id_new]]$mergeCells <-
    self$worksheets[[id_org]]$mergeCells

  invisible(self)
}

workbook_add_sparklines <- function(
    self,
    private,
    sheet = current_sheet(),
    sparklines
) {
  sheet <- private$get_sheet_index(sheet)
  self$worksheets[[sheet]]$add_sparklines(sparklines)
  invisible(self)
}
