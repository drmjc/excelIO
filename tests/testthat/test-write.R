context("Export testing")
require(datasets)
# test_that("input test files exist", {
# 	expect_true(file.exists(system.file("examples", "iris.xls", package="excelIO")))
# })

test_that("tables can be written to xls", {
	write.xls(iris, test1.out<-tempfile(pattern="excelIO", fileext=".xls")); expect_true(file.exists(test1.out))
	write.xls(iris, test2.out<-tempfile(pattern="excelIO", fileext=".xls"), row.names=NA); expect_true(file.exists(test2.out))
	write.xls(iris, test3.out<-tempfile(pattern="excelIO", fileext=".xls"), row.names=NULL); expect_true(file.exists(test3.out))
	write.xls(iris, test4.out<-tempfile(pattern="excelIO", fileext=".xls"), row.names=TRUE); expect_true(file.exists(test4.out))
	write.xls(iris, test5.out<-tempfile(pattern="excelIO", fileext=".xls"), row.names=FALSE); expect_true(file.exists(test5.out))
	write.xls(iris, test6.out<-tempfile(pattern="excelIO", fileext=".xls"), row.names="RowIndex"); expect_true(file.exists(test6.out))
	
	expect_equal(read.xls(test2.out), read.xls(test3.out))
})

test_that("vectors can be written to xls", {
	writeLines.xls(letters, test1.out<-tempfile(pattern="excelIO", fileext=".xls")); expect_true(file.exists(test1.out))
	writeLines.xls(LETTERS, test2.out<-tempfile(pattern="excelIO", fileext=".xls")); expect_true(file.exists(test2.out))
	writeLines.xls(letters, test3.out<-tempfile(pattern="excelIO", fileext=".xls"), names=TRUE); expect_true(file.exists(test3.out))
	writeLines.xls(LETTERS, test4.out<-tempfile(pattern="excelIO", fileext=".xls"), names=TRUE); expect_true(file.exists(test4.out))

	a <- 1:26
	names(a) <- letters
	writeLines.xls(a, test5.out<-tempfile(pattern="excelIO", fileext=".xls"), names=TRUE); expect_true(file.exists(test5.out))
	writeLines.xls(a, test6.out<-tempfile(pattern="excelIO", fileext=".xls"), names=FALSE); expect_true(file.exists(test6.out))
})

test_that("tables can be written to xlsx", {
	write.xlsx(iris, test1.out<-tempfile(pattern="excelIO", fileext=".xlsx")); expect_true(file.exists(test1.out))
	write.xlsx(iris, test2.out<-tempfile(pattern="excelIO", fileext=".xlsx"), row.names=NA); expect_true(file.exists(test2.out))
	write.xlsx(iris, test3.out<-tempfile(pattern="excelIO", fileext=".xlsx"), row.names=NULL); expect_true(file.exists(test3.out))
	write.xlsx(iris, test4.out<-tempfile(pattern="excelIO", fileext=".xlsx"), row.names=TRUE); expect_true(file.exists(test4.out))
	write.xlsx(iris, test5.out<-tempfile(pattern="excelIO", fileext=".xlsx"), row.names=FALSE); expect_true(file.exists(test5.out))
	write.xlsx(iris, test6.out<-tempfile(pattern="excelIO", fileext=".xlsx"), row.names="RowIndex"); expect_true(file.exists(test6.out))
	
	expect_equal(read.xls(test2.out), read.xls(test3.out))
})

test_that("vectors can be written to xlsx", {
	writeLines.xlsx(letters, test1.out<-tempfile(pattern="excelIO", fileext=".xlsx")); expect_true(file.exists(test1.out))
	writeLines.xlsx(LETTERS, test2.out<-tempfile(pattern="excelIO", fileext=".xlsx")); expect_true(file.exists(test2.out))
	writeLines.xlsx(letters, test3.out<-tempfile(pattern="excelIO", fileext=".xlsx"), names=TRUE); expect_true(file.exists(test3.out))
	writeLines.xlsx(LETTERS, test4.out<-tempfile(pattern="excelIO", fileext=".xlsx"), names=TRUE); expect_true(file.exists(test4.out))

	a <- 1:26
	names(a) <- letters
	writeLines.xlsx(a, test5.out<-tempfile(pattern="excelIO", fileext=".xlsx"), names=TRUE); expect_true(file.exists(test5.out))
	writeLines.xlsx(a, test6.out<-tempfile(pattern="excelIO", fileext=".xlsx"), names=FALSE); expect_true(file.exists(test6.out))
})
