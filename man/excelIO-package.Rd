\docType{package}
\name{excelIO-package}
\alias{excelIO}
\alias{excelIO-package}
\title{Fast import and export of Microsoft Excel files}
\description{
  Sometimes it's just easier to read & write Microsoft
  Excel files. This packages allows import and export of
  Microsoft Excel (xls) files. The excellent gdata package
  already provides read.xls, but occasionally the final
  column will containing trailing spaces, so my version
  fixes this. Users must install the perl module
  Spreadsheet::WriteExcel, by John McNamara
}
\details{
  Most of the heavy lifting for this package is from from
  \sQuote{gdata}, and the Spreadsheet::WriteExcel Perl
  module by John McNamara

  \code{\link{read.xls}} imports a worksheet from an Excel
  file.

  \code{\link{write.xls}} writes matrix-like data, or a
  list of matrix-like data into worksheet(s) in an XLS
  file. It also formats the header row, and allows more
  control over how the \code{row.names} are written.
  \code{\link{writeLines.xls}} writes vector's to excel
  files as 1-column tables, or as a 2-column table if the
  names are also written.
}
\author{
  Mark Cowley
}
\references{
  \url{http://search.cpan.org/dist/Spreadsheet-WriteExcel/}
}
\seealso{
  \code{\link{read.xls}} \code{\link{write.xls}}
  \code{\link{writeLines.xls}}
}
\keyword{package}
