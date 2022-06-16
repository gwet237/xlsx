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

#- Write summary.df to a different worksheet named "summaryStats2".
#  Format top and bottom rows and add solid borders to the table.
wb.df2 <- loadWorkbook(file="./data/chap6datasets1a.xlsx")
my.sheet <- createSheet(wb=wb.df2, sheetName="summaryStats2")

MaxRow <- summary.df %>% 
  summarise(across(c(Group,Target),~"All(Max)"),
            across(!c(Group,Target),max))

# Create 3 cell blocks for the 3 parts of the table (top row +
# bottom row and main table with data values)
topRow.cb <- CellBlock(sheet=my.sheet, startRow=5, startColumn=3,
                       noRows = 1, noColumns = 10)
botRow.cb <- CellBlock(sheet=my.sheet, startRow=11, startColumn=3,
                       noRows = 1, noColumns = 10)
data.cb <- CellBlock(sheet=my.sheet, startRow=6, startColumn=3,
                     noRows = 5, noColumns = 10)

# Define the cell styles that we need.
borders.sty <- Border(color="black", 
                      position=c("BOTTOM", "LEFT","TOP","RIGHT"), 
                      pen=c("BORDER_THIN", "BORDER_THIN",
                            "BORDER_THIN","BORDER_THIN"))

topRow.cs <- CellStyle(wb.df2) + 
  Alignment(horizontal="ALIGN_CENTER") +
  Font(wb.df2,isBold=F,isItalic=F,color="yellow",heightInPoints=12)+
  borders.sty
botRow.cs <- CellStyle(wb.df2) + 
  Alignment(horizontal="ALIGN_CENTER") +
  Font(wb.df2, isBold=T,isItalic=F,color="9") + borders.sty
borders.cs <- CellStyle(wb.df2) + borders.sty

top.fill <- Fill(foregroundColor = "blue", 
                 backgroundColor="blue")
bot.fill <- Fill(foregroundColor = "darkred", 
                 backgroundColor="darkred")

CB.setRowData(cellBlock = topRow.cb,x=colnames(summary.df),
              rowIndex=1,rowStyle = topRow.cs)
CB.setRowData(cellBlock = botRow.cb,x=MaxRow,
              rowIndex=1,rowStyle = botRow.cs)
CB.setMatrixData(cellBlock = data.cb,
                 x=as.matrix(summary.df),
                 startRow=1,startColumn = 1,cellStyle=borders.cs )

CB.setFill(topRow.cb, fill=top.fill, rowIndex=rep(1,10),
           colIndex=1:10)
CB.setFill(botRow.cb, fill=bot.fill, rowIndex=rep(1,10),
           colIndex=1:10)
saveWorkbook(wb.df2, file="./data/chap6datasets1b.xlsx")