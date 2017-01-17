context("XLS/XLSX/TSV testing")

test_that("test files exist", {
	expect_true( file.exists(system.file("examples", "xls", "simple.tsv", package="excelIO")) )
	expect_true( file.exists(system.file("examples", "xls", "simple-CRAN-dataframes2xls.xls", package="excelIO")) )
	expect_true( file.exists(system.file("examples", "xls", "simple-CRAN-xlsx.xlsx", package="excelIO")) )
	expect_true( file.exists(system.file("examples", "xls", "simple-Excel2011.xls", package="excelIO")) )
	expect_true( file.exists(system.file("examples", "xls", "simple-Excel2011.xlsx", package="excelIO")) )
	expect_true( file.exists(system.file("examples", "xls", "simple-excelIO.xls", package="excelIO")) )
})

test_that("is.excel.file works", {
	a <- system.file("examples", "xls", "simple.tsv", package="excelIO")
	b <- system.file("examples", "xls", "simple-CRAN-dataframes2xls.xls", package="excelIO")
	c <- system.file("examples", "xls", "simple-CRAN-xlsx.xlsx", package="excelIO")
	d <- system.file("examples", "xls", "simple-Excel2011.xls", package="excelIO")
	e <- system.file("examples", "xls", "simple-Excel2011.xlsx", package="excelIO")
	f <- system.file("examples", "xls", "simple-excelIO.xls", package="excelIO")
	
	expect_false( unname(is.excel.file(a)) )
	expect_true ( unname(is.excel.file(b)) )
	expect_true ( unname(is.excel.file(c)) )
	expect_true ( unname(is.excel.file(d)) )
	expect_true ( unname(is.excel.file(e)) )
	expect_true ( unname(is.excel.file(f)) )
})

test_that("is.xls.file works", {
	a <- system.file("examples", "xls", "simple.tsv", package="excelIO")
	b <- system.file("examples", "xls", "simple-CRAN-dataframes2xls.xls", package="excelIO")
	c <- system.file("examples", "xls", "simple-CRAN-xlsx.xlsx", package="excelIO")
	d <- system.file("examples", "xls", "simple-Excel2011.xls", package="excelIO")
	e <- system.file("examples", "xls", "simple-Excel2011.xlsx", package="excelIO")
	f <- system.file("examples", "xls", "simple-excelIO.xls", package="excelIO")
	
	expect_false( unname(is.xls.file(a)) )
	expect_true ( unname(is.xls.file(b)) )
	expect_false( unname(is.xls.file(c)) )
	expect_true ( unname(is.xls.file(d)) )
	expect_false( unname(is.xls.file(e)) )
	expect_true ( unname(is.xls.file(f)) )
})

test_that("is.xlsx.file works", {
	a <- system.file("examples", "xls", "simple.tsv", package="excelIO")
	b <- system.file("examples", "xls", "simple-CRAN-dataframes2xls.xls", package="excelIO")
	c <- system.file("examples", "xls", "simple-CRAN-xlsx.xlsx", package="excelIO")
	d <- system.file("examples", "xls", "simple-Excel2011.xls", package="excelIO")
	e <- system.file("examples", "xls", "simple-Excel2011.xlsx", package="excelIO")
	f <- system.file("examples", "xls", "simple-excelIO.xls", package="excelIO")
	
	expect_false( unname(is.xlsx.file(a)) )
	expect_false( unname(is.xlsx.file(b)) )
	expect_true ( unname(is.xlsx.file(c)) )
	expect_false( unname(is.xlsx.file(d)) )
	expect_true ( unname(is.xlsx.file(e)) )
	expect_false( unname(is.xlsx.file(f)) )
})

test_that("is.tsv.file works", {
	a <- system.file("examples", "xls", "simple.tsv", package="excelIO")
	b <- system.file("examples", "xls", "simple-CRAN-dataframes2xls.xls", package="excelIO")
	c <- system.file("examples", "xls", "simple-CRAN-xlsx.xlsx", package="excelIO")
	d <- system.file("examples", "xls", "simple-Excel2011.xls", package="excelIO")
	e <- system.file("examples", "xls", "simple-Excel2011.xlsx", package="excelIO")
	f <- system.file("examples", "xls", "simple-excelIO.xls", package="excelIO")
	
	expect_true ( unname(is.tsv.file(a)) )
	expect_false( unname(is.tsv.file(b)) )
	expect_false( unname(is.tsv.file(c)) )
	expect_false( unname(is.tsv.file(d)) )
	expect_false( unname(is.tsv.file(e)) )
	expect_false( unname(is.tsv.file(f)) )
})
