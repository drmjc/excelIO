#' is the file a comma separated text file?
#' 
#' Rather than just checking the file extension, this function checks
#' whether the contents of the file contain tabs.
#' The file paths must exist, they must be ASCII test, and must have at least
#' one comma in one of the lines to qualify. 
#' 
#' @note We could make this more stringent by requiring that all lines contain
#' the same number of comma's, but need to consider how to skip header lines, and
#' column name lines which can contain NCOL-1 columns (if written using write.table),
#' and what to do with rogue comma's in the middle of column
#'
#' @param path a character vector: the path(s) to file(s) which must exist
#' @return a named logical vector of length(path), where names are the paths, 
#' and values are \code{TRUE} if the path is an excel file, \code{FALSE} 
#' otherwise.
#' @author Mark Cowley, 2012-02-01
#' @export
#' @examples
#' \dontrun{
#' is.csv.file("./tmp.xls")
#' is.csv.file("./tmp.tsv")
#' is.csv.file("./tmp.csv")
#' is.csv.file(c("./tmp.xls", "./tmp.tsv", "./tmp.csv"))
#' }
is.csv.file <- function(path) {
	length(path) >= 1 && all(file.exists(path)) || stop("all file paths in 'path' must exist")

	res <- rep(FALSE, length(path))
	names(res) <- path

	for(i in 1:length(path)) {
		file.out <- system( paste("file", path[i]), intern=TRUE, ignore.stderr=TRUE)
		if( any(grepl("ASCII", file.out)) ) {
			cmd <- sprintf("cat '%s' | awk -F'\",\"' '{ print NF }'", path[i])
			tmp <- system(cmd, intern=TRUE, ignore.stderr=TRUE)
			tmp <- as.numeric(tmp)
			res[i] <- any(tmp > 1)
		}
	}
	
	res
}
