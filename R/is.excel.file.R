#' is the file a Microsoft Excel XLS or XLSX file?
#' 
#' Rather than just checking the file extension, this function checks
#' whether the contents of the file are a Microsoft Excel file, using the
#' \code{GNU file} program, thus this will not work on Windows. \code{file}
#' works by checking the file signature to see what file type it is.
#' Note that XLS files can have different file signatures depending where
#' they are saved from/created.
#' 
#' @param path a character vector: the path(s) to file(s) which must exist
#' 
#' @return A \code{logical vector} of \code{length(path)}, where 
#' the values are \code{TRUE} if the path is an XLS, XLSX, or both, \code{FALSE} 
#' otherwise, for \code{\link{is.xls.file}}, \code{\link{is.xlsx.file}}, or 
#' \code{\link{is.excel.file}}, respectively.
#' 
#' @author Mark Cowley, 2012-02-01
#' @seealso \code{\link{is.xls.file}} \code{\link{is.xlsx.file}}
#' @export
#' 
#' @examples
#' \dontrun{
#' # df <- data.frame(A=1:10, B=letters[1:10], C=LETTERS[1:10], D=TRUE, E=FALSE)
#' require(datasets)
#' write.table(iris, "tmp.tsv", quote=FALSE, sep="\t")
#' write.csv(iris, "tmp.csv")
#' require(dataframes2xls)
#' write.xls(iris, "tmp.xls")
#' require(xlsx)
#' write.xlsx(iris, "tmp.xlsx")
#' 
#' is.excel.file("./tmp.xls")
#' is.excel.file("./tmp.xlsx")
#' is.excel.file("./tmp.tsv")
#' is.excel.file("./tmp.csv")
#' is.excel.file(c("./tmp.xls", "./tmp.xlsx", "./tmp.tsv", "./tmp.csv"))
#' }
is.excel.file <- function(path) {
	res <- is.xls.file(path) | is.xlsx.file(path)
	
	res
}
# 2012-03-06
# - expanded path & added shQuote to allow the system command to find unusual filenames.
# 2012-03-08
# - refactored  to is.xls.file and is.xlsx.file
# - removed names so that testthat::expect_true doesn't keep failing

#' @export
#' @rdname is.excel.file
is.xls.file <- function(path) {
	.is.excel.file(path, "xls")
}


#' @export
#' @rdname is.excel.file
is.xlsx.file <- function(path) {
	.is.excel.file(path, "xlsx")
}

#' Is the file and XLS or XLSX file?
#' 
#' This is the private function that does the work for:\cr
#' \code{\link{is.excel.file}}, \code{\link{is.xls.file}}, \code{\link{is.xlsx.file}}
#' @inheritParams is.excel.file
#' @param type one of \dQuote{xls} or \dQuote{xlsx}
#' @return a logical vector
#' @author Mark Cowley, 2012-03-08
#' @noRd
.is.excel.file <- function(path, type=c("xls", "xlsx")) {
	length(path) >= 1 && all(file.exists(path)) || stop("all file paths in 'path' must exist")
	length(type) == 1 || stop("choose one value for type: xls, xlsx")
	res <- rep(FALSE, length(path))
	# names(res) <- path

	path <- normalizePath(path)
	for(i in seq(along=path)) {
		file.out <- system( paste("file -b", shQuote(path[i])), intern=TRUE, ignore.stderr=TRUE)
		res[i] <- 
			if( type == "xlsx") grepl("Zip archive data", file.out) || grepl("Microsoft Excel 2007", file.out) || grepl("Microsoft OOXML", file.out)
			else if( type == "xls" ) grepl("CDF V2 Document", file.out) || grepl("Composite Document File V2 Document", file.out)
	}
	
	res
}
# 2012-07-20
# - get.full.path -> normalizePath
# 2017-01-17
# - updated file -b outputs to work on MacOS Sierra