context("Import testing")

require(datasets)

test_that("input test files exist", {
	expect_true(file.exists(system.file("examples", "iris.xls", package="excelIO")))
})

test_that("xls can be read", {
	iris.xls <- system.file("examples", "iris.xls", package="excelIO")
	a <- read.xls(iris.xls, stringsAsFactors=TRUE)
	expect_is(a, "data.frame")
	expect_equal(dim(a), c(150,5))
	expect_equal(colnames(a), c("Sepal.Length", "Sepal.Width", "Petal.Length", "Petal.Width", "Species"))
	expect_equal(a[,1:4], iris[,1:4])
	expect_equal(a[,5], iris[,5])
	expect_equal(a, iris)
})
