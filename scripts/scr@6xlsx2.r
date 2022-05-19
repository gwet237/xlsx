library(xlsx)
passWD <- "Microsoft365"
wb.df <- loadWorkbook(file="./data/chap6datasets1.xlsx",
                       password = passWD)
my.sheets <- getSheets(wb.df)
print(names(my.sheets))

tit.rows <- getRows(sheet = my.sheets[["summaryStats2"]],
                    rowIndex=1:4)
cells.tit <- getCells(row=tit.rows,colIndex = 3:8)
setCellValue(cells.tit[["1.3"]],
             "Table 1: Mean Ratings by Group and Target, and")
setCellValue(cells.tit[["2.4"]],
             "Group-level standard deviations")
setCellValue(cells.tit[["3.3"]],"(Author: Kilem L. Gwet)")
saveWorkbook(wb.df, file="./data/chap6datasets2.xlsx")