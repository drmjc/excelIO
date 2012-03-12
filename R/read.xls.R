#' Read a Microsoft Excel XLS file into a data frame
#'
#' \code{gdata} provides the wonderful \code{\link[gdata]{read.xls}}, but 
#' frequently if the last column is
#' characters, then there is a single extra space in all words in the last
#' column. This method over-rides \code{gdata}'s, and fixes the last columns. 
#' 
#' This function works translating the named Microsoft Excel file
#' into a temporary .csv or .tab file, using the \code{xls2csv} or \code{xls2tab}
#' Perl script installed as part of the \code{gdata} package.
#' 
#' Caution: In the conversion to csv, strings will be quoted. This
#' can be problem if you are trying to use the \code{comment.char} option
#' of \code{\link{read.table}} since the first character of all lines (including
#' comment lines) will be \dQuote{\"} after conversion.
#' 
#' @note This extra space is probably due to different line endings.
#' 
#' @param xls path to the Microsoft Excel file.  Supports \dQuote{http://},
#'          \dQuote{https://}, and \dQuote{ftp://} URLs.
#' @param sheet number of the sheet within the Excel file from which data are
#'          to be read
#' @param verbose logical flag indicating whether details should be printed as
#'          the file is processed.
#' @param \dots additional arguments to read.table. The defaults for \code{\link{read.csv}}() are used.
#'  hint: try \code{check.names=FALSE}
#' @return a \code{data.frame}
#' @author Mark Cowley, 2009-01-22
#' @export
#' @importFrom gdata read.xls trim
read.xls <- function(xls, sheet = 1, verbose=FALSE, ...) {
	!missing(xls) || stop("xls must be the path to an excel file")
	is.excel.file(xls) || stop("xls is not a Microsoft Excel file. see ?is.excel.file")
	
	res <- gdata::read.xls(xls=xls, sheet=sheet, verbose=verbose, ...)
	NCOL <- ncol(res)
	if( class(res[, NCOL]) == "character" ) {
		res[, NCOL] <- gdata::trim(res[, NCOL])
	}
	else if( class(res[, NCOL]) == "factor" ) {
		levels(res[, NCOL]) <- gdata::trim(levels(res[, NCOL]))
	}
	
	res
}
