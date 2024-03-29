% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/class-workbook-wrappers.R
\name{wb_add_comment}
\alias{wb_add_comment}
\alias{wb_get_comment}
\alias{wb_remove_comment}
\title{Add comment to worksheet}
\usage{
wb_add_comment(wb, sheet = current_sheet(), dims = "A1", comment, ...)

wb_get_comment(wb, sheet = current_sheet(), dims = NULL)

wb_remove_comment(wb, sheet = current_sheet(), dims = "A1", ...)
}
\arguments{
\item{wb}{A workbook object}

\item{sheet}{A worksheet of the workbook}

\item{dims}{Optional row and column as spreadsheet dimension, e.g. "A1"}

\item{comment}{A comment to apply to \code{dims} created by \code{\link[=wb_comment]{wb_comment()}}, a string or a \code{\link[=fmt_txt]{fmt_txt()}} object}

\item{...}{additional arguments}
}
\value{
The Workbook object, invisibly.
}
\description{
Add comment to worksheet
}
\details{
If applying a \code{comment} with a string, it will use \code{\link[=wb_comment]{wb_comment()}} default values. If additional background colors are applied, RGB colors should be provided, either as hex code or with builtin R colors. The alpha channel is ignored.
}
\examples{
wb <- wb_workbook()
wb$add_worksheet("Sheet 1")
# add a comment without author
c1 <- wb_comment(text = "this is a comment", author = "")
wb$add_comment(dims = "B10", comment = c1)
#' # Remove comment
wb$remove_comment(sheet = "Sheet 1", dims = "B10")
# Write another comment with author information
c2 <- wb_comment(text = "this is another comment", author = "Marco Polo", visible = TRUE)
wb$add_comment(sheet = 1, dims = "C10", comment = c2)
# Works with formatted text also.
formatted_text <- fmt_txt("bar", underline = TRUE)
wb$add_comment(dims = "B5", comment = formatted_text)
# With background color
wb$add_comment(dims = "B7", comment = formatted_text, color = wb_color("green"))
# With background image. File extension must be png or jpeg, not jpg?
tmp <- tempfile(fileext = ".png")
png(file = tmp, bg = "transparent")
plot(1:10)
rect(1, 5, 3, 7, col = "white")
dev.off()

c1 <- wb_comment(text = "this is a comment", author = "", visible = TRUE)
wb$add_comment(dims = "B12", comment = c1, file = tmp)
}
\seealso{
\code{\link[=wb_comment]{wb_comment()}}, \code{\link[=wb_add_thread]{wb_add_thread()}}
}
\keyword{comments}
