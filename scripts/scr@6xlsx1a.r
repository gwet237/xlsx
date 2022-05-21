library(tidyverse)
library(xlsx)
passWD <- "Microsoft365"
wb.df1 <- loadWorkbook(file="./data/chap6datasets.xlsx",
                       password = passWD)
sht.names <- names(getSheets(wb.df1))

#-- Reading the password-protected Excel workbook -
irrData.df <- read.xlsx(file="./data/chap6datasets.xlsx", 
                        sheetName = "IrrData",
                        rowIndex = c(2:17),colIndex = c(19:24),
                        password = passWD)

#-- Compute the summary table
summary.df <- as_tibble(irrData.df) %>%
  group_by(Group,Target) %>%
  summarise(across(c(J1,J2,J3,J4),mean)) %>%
  mutate(across(c(J1,J2,J3,J4),max,.names="gMax.{.col}")) %>%
  ungroup() %>%
  mutate(across(!c(Group,Target),round,3))
print(summary.df)

#-- Write summary table to chap6datasets1a.xlsx (no password)
write.xlsx(x = irrData.df, 
           file="./data/chap6datasets1a.xlsx", 
           sheetName = "Input Data", row.names = FALSE)
write.xlsx(x = as.data.frame(summary.df), 
           file="./data/chap6datasets1a.xlsx", 
           sheetName = "summaryStats1",append = TRUE,
           row.names = FALSE)

#- Write the summary.df at a specific worksheet location
wb.df1a <- loadWorkbook(file="./data/chap6datasets1a.xlsx")
my.sheet <- createSheet(wb=wb.df1a, sheetName="summaryStats2")

max.row<-summary.df %>% 
  summarise(across(c(Group,Target),~"Max Value"),
            across(!c(Group,Target),max))
summary.df %>%
  mutate(Target=as.character(Target))%>%
  bind_rows(max.row) -> summary1.df

addDataFrame(x = as.data.frame(summary1.df), sheet = my.sheet,
             col.names = TRUE,row.names = FALSE,
             startRow=5, startColumn=3)
saveWorkbook(wb.df1a, file="./data/chap6datasets1a.xlsx")
