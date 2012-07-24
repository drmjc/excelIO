#' is the file a tab separated text file?
#' 
#' Rather than just checking the file extension, this function checks
#' whether the contents of the file contain tabs.
#' The file paths must exist, they must be ASCII test, and must have at least
#' one tab in one of the lines to qualify. 
#' 
#' @note We could make this more stringent by requiring that all lines contain
#' the same number of tabs, but need to consider how to skip header lines, and
#' column name lines which can contain NCOL-1 columns (if written using write.table)
#'
#' @param path a character vector: the path(s) to file(s) which must exist
#' @return a logical vector of length(path), where
#' the values are \code{TRUE} if the path is an excel file, \code{FALSE} 
#' otherwise.
#' @author Mark Cowley, 2012-02-01
#' @export
#' @examples
#' \dontrun{
#' is.tsv.file("./tmp.xls")
#' is.tsv.file("./tmp.tsv")
#' is.tsv.file("./tmp.csv")
#' is.tsv.file(c("./tmp.xls", "./tmp.tsv", "./tmp.csv"))
#' }
is.tsv.file <- function(path) {
	length(path) >= 1 && all(file.exists(path)) || stop("all file paths in 'path' must exist")
	path <- normalizePath(path)
	
	res <- rep(FALSE, length(path))
	# names(res) <- path

	for(i in 1:length(path)) {
		file.out <- system( paste("file -b", path[i]), intern=TRUE, ignore.stderr=TRUE)
		if( grepl("ASCII", file.out) ) {
			cmd <- sprintf("cat %s | awk -F'\\\t' '{ print NF }'", shQuote(path[i]))
			tmp <- system(cmd, intern=TRUE, ignore.stderr=TRUE)
			tmp <- as.numeric(tmp)
			res[i] <- any(tmp > 1)
		}
	}
	
	res
}
# 2012-03-06
# - expanded path & added shQuote to allow the system command to find unusual filenames.
# - removed names due to causing testthat::expect_true to fail
# 2012-07-20
# - get.full.path -> normalizePath
