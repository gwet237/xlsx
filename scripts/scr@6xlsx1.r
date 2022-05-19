library(tidyverse)
library(xlsx)
passWD <- "Microsoft365"
wb.df1 <- loadWorkbook(file="./data/chap6datasets.xlsx",
                       password = passWD)
sht.names <- names(getSheets(wb.df1))
print(sht.names)

#-- Reading the password-protected Excel workbook -
irrData.df <- read.xlsx(file="./data/chap6datasets.xlsx", 
                        sheetName = "IrrData",
                        rowIndex = c(2:17),colIndex = c(19:24),
                        password = passWD)

#-- Compute the summary table
summary.df <- as_tibble(irrData.df) %>%
  group_by(Group,Target) %>%
  summarise(across(c(J1,J2,J3,J4),mean)) %>%
  mutate(across(c(J1,J2,J3,J4),sd,.names="gsd.{.col}")) %>%
  ungroup() %>%
  mutate(across(!c(Group,Target),round,3))
print(summary.df)

#-- Write the summary table to the input workbook
write.xlsx(x = as.data.frame(summary.df), 
           file="./data/chap6datasets1.xlsx", 
           sheetName = "summaryStats1",append = TRUE,
           row.names = FALSE, password = passWD)

#- Write the summary.df at a specific worksheet location
wb.df1 <- loadWorkbook(file="./data/chap6datasets1.xlsx",
                       password = "Microsoft365")
my.sheet <- createSheet(wb=wb.df1, sheetName="summaryStats2")
addDataFrame(x = as.data.frame(summary.df), sheet = my.sheet,
             col.names = TRUE,row.names = FALSE,
             startRow=5, startColumn=3)
saveWorkbook(wb.df1, file="./data/chap6datasets.xlsx",
             password = passWD)

#-- Delete the summaryStats2 worksheet if necessary
removeSheet(wb=wb.df1, sheetName = "summaryStats1")
saveWorkbook(wb.df1, file="./data/chap6datasets.xlsx",
             password = passWD)
