#' Write to Microsoft Excel XLS file
#' 
#' Write a \code{matrix} or \code{data.frame} to a genuine Microsoft Excel file.
#' The first row is bold, italics, dashed underline. Columns are centred.
#' 
#' Uses a perl module called \code{Spreadsheet::WriteExcel}, by John McNamara,
#' a perl program called \code{tab2xls}, based on a program of the same name also by
#' JM.
#' This R code borrows the structure from \code{gdata}'s excellent 
#' \code{\link[gdata]{read.xls}} function.
#' 
#' @note Users must install the Spreadsheet::WriteExcel Perl module.\cr
#' For OSX users, this involves:\cr
#' sudo perl -e shell -MCPAN\cr
#' install Spreadsheet::WriteExcel\cr
#' - or -\cr
#' sudo cpan\cr
#' ### accept the defaults, and set up which country you are in.\cr
#' install CPAN\cr
#' reload cpan\cr
#' install Spreadsheet::WriteExcel\cr
#' 
#' @param x either a data.frame, or a list of data.frames. if a list of
#'   data.frames, then each data.frame will become a worksheet in the resulting
#'   workbook.
#' @param xls the xls file to be created.
#' @param verbose TRUE/FALSE
#' @param row.names DIFFERENT TO \code{\link{write.table}}. Can be either \code{TRUE}, \code{FALSE}, \code{NULL},
#'   \code{NA}, or a \code{character(1)}. if \code{TRUE} then row.names will be written out, with a
#'   blank column header, if \code{FALSE}, then no row names will be written out, if
#'   \code{NULL} or \code{NA}, then rownames will be written out with no column name (just
#'   like \code{write.table(row.names=TRUE)} would do) if \code{row.names} is a character
#'   vector of length 1, then write the \code{row.names} out, and use use the supplied
#'   word as the column name for the rownames else error
#' @param col.names logical: If \code{TRUE} then write out the column names.
#' @param \dots additional arguments passed to \code{\link{write.delim}}
#' @return none. writes an excel 'xls' file.
#' @author Mark Cowley, 2009-01-20
#' @seealso \code{\link{write.table}} \code{\link{read.delim}} \code{\link[gdata]{read.xls}}
#' @export
#' @examples
#' f <- tempfile(fileext=".xls")
#' write.xls(iris, f)
#' x <- list(A=iris[1:5, ], B=iris)
#' write.xls(x, f)
#' 
write.xls <- function(x, xls, verbose=FALSE, row.names=FALSE, col.names=TRUE, ...) {
	!missing(x) || stop("Must specify x")
	!missing(xls) && file.exists(dirname(xls)) || stop(sprintf("Cannot print to %s because %s does not exist.\n", xls, dirname(xls)))

	if( is.vector(x) && !is.list(x) ) {
		writeLines.xls(x, xls, row.names==TRUE)
		return("done")
	}
		
	if( is.matrix(x) || is.data.frame(x) ) {
		x <- list(x)
		names(x) <- "Sheet1"
	}
	else {
		if(is.null(names(x))) {
			names(x) <- paste("Sheet", 1:length(x), sep="")
		}
	}
	
	# check that all tables are small enough to fit in an XLS
	for(elem in seq(along=x)) {
		if( (nrow(x[[elem]]) + as.numeric(col.names)) > 65536 ) {
			warning("data exceeds the 65,536 maximum number of rows allowed in an excel file")
			x[[elem]] <- x[[elem]][1:(65536 - col.names), ]
		}
		
	}

	###
	tab2xls <- file.path(.path.package('excelIO'), 'bin', 'tab2xls')
	file.exists(tab2xls) || stop("Couldn't find tab2xls, thus can't create Excel files. Are you sure you have excelIO installed?")
	###

	###
	# export each table in x as a tab delimited file into a tmpfile.
	# The tmpfile names will eventually become the worksheet names, so set them accordingly
	# also, worksheet names have to be <= 31 chars in length (at least in excel 2004 on mac)
	# 
	txtfiles <- rep(NA, length(x))
	dir <- tempdir()
	for(i in 1:length(x)) {
		f <- paste(substring(names(x)[i], 1, 31), "xls", sep=".")
		txtfiles[i] <- file.path(dir, f)
		write.delim(x[[i]], txtfiles[i], row.names=row.names, col.names=col.names, ...)
	}
	# format these file names suitable for a system call
	xls.cmd <- shQuote(path.expand(xls))
	txtfiles.cmd <- paste(shQuote(txtfiles), collapse=" ")
	#
	###

	###
	# execution command
	cmd <- paste("perl", tab2xls, txtfiles.cmd, xls.cmd, sep=" ")
	#
	###

	###
	# do the translation
	if(verbose)  cat("Executing ", cmd, "... \n")
	#
	results <- system(cmd, intern=!verbose)
	#
	if (verbose) cat("done.\n")
	#
	###

	if( any(!file.remove(txtfiles)) )
		stop("Warning, some temp files were not deleted.\n")

}
# CHANGELOG
# 2013-04-15: bug fix, when x is a list.
