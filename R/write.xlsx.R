#' Write to Microsoft Excel XLSX file
#' 
#' Write a \code{matrix} or \code{data.frame} to a genuine Microsoft Excel XLSX file.
#' The first row is bold, italics, dashed underline. Columns are centred.
#' 
#' Uses a perl module called \code{Excel::Writer::XLSX}, by John McNamara,
#' a perl program called \code{tab2xlsx}, based on a program of the same name also by
#' JM.
#' This R code borrows the structure from \code{gdata}'s excellent 
#' \code{\link[gdata]{read.xls}} function.
#' 
#' @note Installation
#' It's not that widely tested, but this package does include the necessary perl libraries to
#' write XLSX files. If you're struggling though, install the `Excel::Writer::XLSX` Perl module
#' via CPAN or similar. For OSX users, this involves:\cr
#' sudo perl -e shell -MCPAN\cr
#' install Excel::Writer::XLSX\cr
#' 
#' @param x either a \code{data.frame}, or a list of \code{data.frame}s. if a list of
#'   \code{data.frame}s, then each \code{data.frame} will become a worksheet in the resulting
#'   workbook.
#' @param xlsx the xlsx file to be created.
#' @param verbose logical: print messages?
#' @param row.names DIFFERENT TO \code{\link{write.table}}. Can be either \code{TRUE}, \code{FALSE}, \code{NULL},
#'   \code{NA}, or a \code{character(1)}. if \code{TRUE} then row.names will be written out, with a
#'   blank column header, if \code{FALSE}, then no row names will be written out, if
#'   \code{NULL} or \code{NA}, then rownames will be written out with no column name (just
#'   like \code{write.table(row.names=TRUE)} would do) if \code{row.names} is a character
#'   vector of length 1, then write the \code{row.names} out, and use use the supplied
#'   word as the column name for the rownames else error
#' @param col.names logical: If \code{TRUE} then write out the column names.
#' @param \dots additional arguments passed to \code{\link{write.delim}}
#' @return none. writes an excel 'xlsx' file.
#' @author Mark Cowley, 2009-01-20. Updated 2016-04-14 to write XLSX files.
#' @seealso \code{\link{write.xls}} \code{\link{write.table}} \code{\link{read.delim}} \code{\link[gdata]{read.xls}}
#' @export
#' @examples
#' f <- tempfile(fileext=".xlsx")
#' write.xlsx(iris, f)
#' x <- list(A=iris[1:5, ], B=iris)
#' write.xlsx(x, f)
#' 
write.xlsx <- function(x, xlsx, verbose=FALSE, row.names=FALSE, col.names=TRUE, ...) {
	!missing(x) || stop("Must specify x")
	!missing(xlsx) && file.exists(dirname(xlsx)) || stop(sprintf("Cannot print to %s because %s does not exist.\n", xlsx, dirname(xlsx)))

	if( is.vector(x) && !is.list(x) ) {
		writeLines.xlsx(x, xlsx, row.names==TRUE)
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
	
	# check that all tables are small enough to fit in an xlsx
	for(elem in seq(along=x)) {
		if( (nrow(x[[elem]]) + as.numeric(col.names)) > 65536 ) {
			warning("data exceeds the 65,536 maximum number of rows allowed in an excel file")
			x[[elem]] <- x[[elem]][1:(65536 - col.names), ]
		}
		
	}

	### this path needs to work for testthat (ie locally) and once installed.
	tab2xlsx <- c(
		file.path(path.package('excelIO'), 'bin', 'tab2xlsx'),
		file.path(path.package('excelIO'), 'inst', 'bin', 'tab2xlsx')
	)
	tab2xlsx <- tab2xlsx[file.exists(tab2xlsx)]
	length(tab2xlsx) == 1 || stop(paste("Couldn't find tab2xlsx, at ", tab2xlsx))
	###

	###
	# export each table in x as a tab delimited file into a tmpfile.
	# The tmpfile names will eventually become the worksheet names, so set them accordingly
	# also, worksheet names have to be <= 31 chars in length (at least in excel 2004 on mac)
	# 
	txtfiles <- rep(NA, length(x))
	dir <- tempdir()
	for(i in 1:length(x)) {
		f <- paste(substring(names(x)[i], 1, 31), "xlsx", sep=".")
		txtfiles[i] <- file.path(dir, f)
		write.delim(x[[i]], txtfiles[i], row.names=row.names, col.names=col.names, ...)
	}
	# format these file names suitable for a system call
	xlsx.cmd <- shQuote(path.expand(xlsx))
	txtfiles.cmd <- paste(shQuote(txtfiles), collapse=" ")
	#
	###

	###
	# execution command
	cmd <- paste("perl", tab2xlsx, txtfiles.cmd, xlsx.cmd, sep=" ")
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
# 2016-04-14: clone from write.xls, and updated to use Excel::Writer::XLSX
# 2017-01-17: update tab2xlsx path to let testthat work locally