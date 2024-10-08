% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/class-workbook-wrappers.R
\name{wb_add_chartsheet}
\alias{wb_add_chartsheet}
\title{Add a chartsheet to a workbook}
\usage{
wb_add_chartsheet(
  wb,
  sheet = next_sheet(),
  tab_color = NULL,
  zoom = 100,
  visible = c("true", "false", "hidden", "visible", "veryhidden"),
  ...
)
}
\arguments{
\item{wb}{A Workbook object to attach the new chartsheet}

\item{sheet}{A name for the new chartsheet}

\item{tab_color}{Color of the sheet tab. A  \code{\link[=wb_color]{wb_color()}},  a valid color (belonging to
\code{grDevices::colors()}) or a valid hex color beginning with "#".}

\item{zoom}{The sheet zoom level, a numeric between 10 and 400 as a
percentage. (A zoom value smaller than 10 will default to 10.)}

\item{visible}{If \code{FALSE}, sheet is hidden else visible.}

\item{...}{Additional arguments}
}
\description{
A chartsheet is a special type of sheet that handles charts output. You must
add a chart to the sheet. Otherwise, this will break the workbook.
}
\seealso{
\code{\link[=wb_add_mschart]{wb_add_mschart()}}

Other workbook wrappers: 
\code{\link{base_font-wb}},
\code{\link{col_widths-wb}},
\code{\link{creators-wb}},
\code{\link{grouping-wb}},
\code{\link{row_heights-wb}},
\code{\link{wb_add_data}()},
\code{\link{wb_add_data_table}()},
\code{\link{wb_add_formula}()},
\code{\link{wb_add_hyperlink}()},
\code{\link{wb_add_pivot_table}()},
\code{\link{wb_add_slicer}()},
\code{\link{wb_add_worksheet}()},
\code{\link{wb_base_colors}},
\code{\link{wb_clone_worksheet}()},
\code{\link{wb_copy_cells}()},
\code{\link{wb_freeze_pane}()},
\code{\link{wb_merge_cells}()},
\code{\link{wb_save}()},
\code{\link{wb_set_last_modified_by}()},
\code{\link{wb_workbook}()}
}
\concept{workbook wrappers}
