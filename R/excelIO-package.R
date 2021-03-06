#' Fast import and export of Microsoft Excel files
#'
#' Sometimes it's just easier to read & write Microsoft
#' Excel files. This packages allows import and export of Microsoft
#' Excel (xls, xlsx) files. The excellent gdata package already provides
#' \code{read.xls}, but occasionally the final column will containing trailing
#' spaces, so my version fixes this.
#' 
#' Reading excel files is powered by the \sQuote{gdata} package.
#' Writing xls files is powered by the Spreadsheet::WriteExcel Perl module by John McNamara,
#' and writing xlsx files is powered by the Excel::Writer::XLSX Perl module by John McNamara.
#'
#' @section Dependencies:
#' It's not that widely tested, but the necessary perl modules are installed within this package,
#' and my version of R appears to be loading these libraries from the package. If it's not working
#' then try installing 
#' Spreadsheet::WriteExcel and Excel::Writer::XLSX system-wide; to do this, see \code{\link{write.xls}}.
#' 
#' @section Key functions:
#' \code{\link{read.xls}} imports a worksheet from an Excel file.
#' 
#' \code{\link{write.xls}}, \code{\link{write.xlsx}} writes matrix-like data, or a list of matrix-like
#' data into worksheet(s) in an XLS file. It also formats the header row, and
#' allows more control over how the \code{row.names} are written.
#' \code{\link{writeLines.xls}} writes vector's to excel files as 1-column
#' tables, or as a 2-column table if the names are also written.
#' 
#' @name excelIO-package
#' @aliases excelIO-package excelIO
#' @docType package
#' @author Mark Cowley
#' @seealso \code{\link{read.xls}} \code{\link{write.xls}} \code{\link{writeLines.xls}}
#' @references \url{http://search.cpan.org/dist/Spreadsheet-WriteExcel/}
#' @keywords package
#' 
NULL
