library(xlsx)
my.wb <- createWorkbook()
my.sheet  <- createSheet(wb=my.wb, sheetName="CellBlock")

my.cb <- CellBlock(sheet=my.sheet, startRow=3, startColumn=2, 
                   noRows = 17, noColumns =  10)
CB.setColData(cellBlock = my.cb, x=1:17, colIndex = 1)
CB.setRowData(cellBlock = my.cb, x=1:10, rowIndex = 1)

# add a matrix to the range of cells and format it
my.cs <- CellStyle(my.wb) + 
  DataFormat("#,##0.000;[yellow]-#,##0.000")
x.mat  <- matrix(rnorm(10*6,2.5,3.7), nrow=10)
CB.setMatrixData(cellBlock = my.cb, x=x.mat, startRow = 4, 
                 startColumn = 3, cellStyle = my.cs)
# add style to the matrix
my.fill <- Fill(foregroundColor = "blue", backgroundColor="blue")
neg.coord <- which(x.mat < 0, arr.ind=TRUE)
CB.setFill(my.cb, fill=my.fill, rowIndex = neg.coord[,1]+3, 
           colIndex=neg.coord[,2]+2)
saveWorkbook(wb=my.wb, file="./data/xlsx formattings.xlsx")