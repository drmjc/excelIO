context("Export testing")
require(datasets)
# test_that("input test files exist", {
# 	expect_true(file.exists(system.file("examples", "iris.xls", package="excelIO")))
# })

test_that("tables can be written to xls", {
	expect_output(write.xls(iris, test1.out<-paste(tempfile(pattern="excelIO"), ".xls", sep="")), "^$")
	expect_output(write.xls(iris, test2.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), row.names=NA), "^$")
	expect_output(write.xls(iris, test3.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), row.names=NULL), "^$")
	expect_output(write.xls(iris, test4.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), row.names=TRUE), "^$")
	expect_output(write.xls(iris, test5.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), row.names=FALSE), "^$")
	expect_output(write.xls(iris, test6.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), row.names="RowIndex"), "^$")
	
	expect_equal(read.xls(test2.out), read.xls(test3.out))
})

test_that("vectors can be written to xls", {
	expect_output(writeLines.xls(letters, test1.out<-paste(tempfile(pattern="excelIO"), ".xls", sep="")), "^$")
	expect_output(writeLines.xls(LETTERS, test2.out<-paste(tempfile(pattern="excelIO"), ".xls", sep="")), "^$")

	expect_output(writeLines.xls(letters, test3.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), names=TRUE), "^$")
	expect_output(writeLines.xls(LETTERS, test4.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), names=TRUE), "^$")

	a <- 1:26
	names(a) <- letters
	expect_output(writeLines.xls(a, test5.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), names=TRUE), "^$")
	expect_output(writeLines.xls(a, test6.out<-paste(tempfile(pattern="excelIO"), ".xls", sep=""), names=FALSE), "^$")
})

test_that("tables can be written to xlsx", {
	expect_output(write.xlsx(iris, test1.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep="")), "^$")
	expect_output(write.xlsx(iris, test2.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), row.names=NA), "^$")
	expect_output(write.xlsx(iris, test3.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), row.names=NULL), "^$")
	expect_output(write.xlsx(iris, test4.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), row.names=TRUE), "^$")
	expect_output(write.xlsx(iris, test5.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), row.names=FALSE), "^$")
	expect_output(write.xlsx(iris, test6.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), row.names="RowIndex"), "^$")
	
	expect_equal(read.xls(test2.out), read.xls(test3.out))
})

test_that("vectors can be written to xlsx", {
	expect_output(writeLines.xlsx(letters, test1.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep="")), "^$")
	expect_output(writeLines.xlsx(LETTERS, test2.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep="")), "^$")

	expect_output(writeLines.xlsx(letters, test3.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), names=TRUE), "^$")
	expect_output(writeLines.xlsx(LETTERS, test4.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), names=TRUE), "^$")

	a <- 1:26
	names(a) <- letters
	expect_output(writeLines.xlsx(a, test5.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), names=TRUE), "^$")
	expect_output(writeLines.xlsx(a, test6.out<-paste(tempfile(pattern="excelIO"), ".xlsx", sep=""), names=FALSE), "^$")
})

