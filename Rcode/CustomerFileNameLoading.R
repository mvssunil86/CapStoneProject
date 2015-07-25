library(xlsx)
#Set working directory 
directory  = "C:/Users/surya/Documents/Files"
setwd(directory)
# Get list of file names into vector 
FileName <- list.files()
# Count no of files 
FilesNo <- length(FileName)
# Create Empty data frame
CustomerList <- data.frame(CustomerName = character() , FileName = character())
# Loop over data Frame
for (i in 1:FilesNo ) {
  df<- read.xlsx(FileName[i], sheetName="Sheet1")
  Colindex <-  grep("CustomerName", colnames(df)) 
  etl <-data.frame( CustomerName = df[,Colindex] ,FileName = FileName[i])
  CustomerList <- rbind(CustomerList,etl)
}
# Dump Customername and FileName to MasterList.Xlsx
library(XLConnect)
wb <- loadWorkbook("MasterList.xlsx", create = TRUE)
createSheet(wb,name="CustomerList")
writeWorksheet(wb,CustomerList,sheet="CustomerList")
saveWorkbook(wb)



