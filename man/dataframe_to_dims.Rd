% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/wb_functions.R
\name{dataframe_to_dims}
\alias{dataframe_to_dims}
\title{Create dimensions from dataframe}
\usage{
dataframe_to_dims(df)
}
\arguments{
\item{df}{dataframe with spreadsheet columns and rows}
}
\description{
Use \code{\link[=wb_dims]{wb_dims()}}
}
\examples{
 df <- dims_to_dataframe("A1:D5;F1:F6;D8", fill = TRUE)
 dataframe_to_dims(df)
}
\keyword{internal}
