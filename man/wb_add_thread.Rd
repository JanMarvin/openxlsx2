% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/class-workbook-wrappers.R
\name{wb_add_thread}
\alias{wb_add_thread}
\alias{wb_get_thread}
\title{Add threaded comments to a cell in a worksheet}
\usage{
wb_add_thread(
  wb,
  sheet = current_sheet(),
  dims = "A1",
  comment = NULL,
  person_id,
  reply = FALSE,
  resolve = FALSE
)

wb_get_thread(wb, sheet = current_sheet(), dims = NULL)
}
\arguments{
\item{wb}{A workbook}

\item{sheet}{A worksheet}

\item{dims}{A cell}

\item{comment}{The text to add, a character vector.}

\item{person_id}{the person Id this should be added. The default is
\code{getOption("openxlsx2.thread_id")} if set.}

\item{reply}{Is the comment a reply? (default \code{FALSE})}

\item{resolve}{Should the comment be resolved? (default \code{FALSE})}
}
\description{
These functions allow adding thread comments to spreadsheets.
This is not yet supported by all spreadsheet software. A threaded comment must
be tied to a person created by \code{\link[=wb_add_person]{wb_add_person()}}.
}
\details{
If a threaded comment is added, it needs a person attached to it.
The default is to create a person with provider id \code{"None"}.
Other providers are possible with specific values for \code{id} and \code{user_id}.
If you require the following, create a workbook via spreadsheet software load
it and get the values with \code{\link[=wb_get_person]{wb_get_person()}}
}
\examples{
wb <- wb_workbook()$add_worksheet()
# Add a person to the workbook.
wb$add_person(name = "someone who likes to edit workbooks")

pid <- wb$get_person(name = "someone who likes to edit workbooks")$id

# write a comment to a thread, reply to one and solve some
wb <- wb_add_thread(wb, dims = "A1", comment = "wow it works!", person_id = pid)
wb <- wb_add_thread(wb, dims = "A2", comment = "indeed", person_id = pid, resolve = TRUE)
wb <- wb_add_thread(wb, dims = "A1", comment = "so cool", person_id = pid, reply = TRUE)
}
\seealso{
\code{\link[=wb_add_comment]{wb_add_comment()}} \code{\link{person-wb}}

Other worksheet content functions: 
\code{\link{col_widths-wb}},
\code{\link{filter-wb}},
\code{\link{grouping-wb}},
\code{\link{named_region-wb}},
\code{\link{row_heights-wb}},
\code{\link{wb_add_conditional_formatting}()},
\code{\link{wb_add_data}()},
\code{\link{wb_add_data_table}()},
\code{\link{wb_add_formula}()},
\code{\link{wb_add_hyperlink}()},
\code{\link{wb_add_pivot_table}()},
\code{\link{wb_add_slicer}()},
\code{\link{wb_freeze_pane}()},
\code{\link{wb_merge_cells}()}
}
\concept{worksheet content functions}
