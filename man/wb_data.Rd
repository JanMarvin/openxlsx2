% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/read.R
\name{wb_data}
\alias{wb_data}
\alias{[.wb_data}
\title{Add the \code{wb_data} attribute to a data frame in a worksheet}
\usage{
wb_data(wb, sheet = current_sheet(), dims, ...)

\method{[}{wb_data}(
  x,
  i,
  j,
  drop = !((missing(j) && length(i) > 1) || (!missing(i) && length(j) > 1))
)
}
\arguments{
\item{wb}{a workbook}

\item{sheet}{a sheet in the workbook either name or index}

\item{dims}{the dimensions}

\item{...}{additional arguments for \code{wb_to_df()}. Be aware that not every
argument is valid.}

\item{x}{x}

\item{i}{i}

\item{j}{j}

\item{drop}{drop}
}
\value{
A data frame of class \code{wb_data}.
}
\description{
provide wb_data object as mschart input
}
\examples{
 wb <- wb_workbook()
 wb <- wb_add_worksheet(wb)
 wb <- wb_add_data(wb, x = mtcars, dims = "B2")

 wb_data(wb, 1, dims = "B2:E6")
}
\seealso{
\code{\link[=wb_to_df]{wb_to_df()}} \code{\link[=wb_add_mschart]{wb_add_mschart()}}, \code{\link[=wb_add_pivot_table]{wb_add_pivot_table()}}
}
