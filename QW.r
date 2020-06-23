QW<-function(week){
  
  Address=paste("U:/001 PERSONAL FOLDERS/Weilun/Weekly Qpseak Report/2020/Week", week) 
  
  setwd(Address)
  filename<-list.files()
  
  Qspeak<-do.call(rbind, lapply(filename, read.csv, header=TRUE, sep="|"))
  
  library(XLConnect)
  
  filename<-paste("AME Qspeak Data Week ", week, ".xlsx", sep="")
  
  wb<-loadWorkbook(filename, create=TRUE)
  sheetname<-paste("Week", week)
  createSheet(wb, sheetname)
  writeWorksheet(wb, Qspeak, sheet=sheetname)
  saveWorkbook(wb)

}
