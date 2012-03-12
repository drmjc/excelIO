#' Write a vector out to an xls file. 
#' 
#' Write a vector out to an xls file. If the vector has names, then optionally
#' write out a two column xls file with names in 1st col, data in 2nd.
#' See the help for \code{\link{write.xls}} for the configuration required.
#' 
#' @param x	a vector
#' @param file the path to an output file
#' @param names logical: if \code{TRUE}, then write out a 2 column table of names then value, 
#'    if \code{FALSE}, then write out just the values in \code{x}.
#' @return nothing.
#' @author Mark Cowley, 2009-02-03
#' @seealso \code{\link{write.xls}}
#' @export
writeLines.xls <- function(x, file, names=TRUE) {
	res <- as.matrix(x, ncol=1)
	
	if( is.null(names) || is.na(names) ) {
		col.names <- FALSE
	}
	else if( is.logical(names) && names && !is.null(names(x)) ) {
		res <- cbind(names=names(x), data=x)
		col.names <- TRUE
	}
	else if ( is.logical(names) && !names ) {
		# do nothing
		res <- as.matrix(x, ncol=1)
		col.names <- FALSE
	}
	else if( is.character(names) && length(names) == 1 ) {
		res <- cbind(names=names(x), data=x)
		colnames(res)[1] <- names
		col.names <- TRUE
	}
	else {
		col.names <- FALSE
	}

	write.xls(res, file, row.names=FALSE, col.names=col.names)
}
