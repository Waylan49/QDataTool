library(XLConnect)
library(tidyverse)

#Lewis's i-Return Receiving File  -------------1
setwd("C:/Users/Weilun_Chiu/Documents/3PAR/iReturn")
filename<-list.files()
datalist<-lapply(filename,read.csv)
lewis_data<-do.call(rbind, datalist)

myda<-lewis_data[, c("DISP.LOCATION", "DATE.SENT", "PART.SERNO")]
myda$DATE.SENT<-as.POSIXct(strptime(myda$DATE.SENT, "%m/%d/%Y %H:%M"))
myda<-myda %>% arrange(desc(DATE.SENT))
myda<-myda %>% distinct(PART.SERNO, .keep_all = TRUE)
mydaF<-data.frame(myda$PART.SERNO, myda$DISP.LOCATION)
mydaF$myda.PART.SERNO<-as.character(mydaF$myda.PART.SERNO)
mydaF$myda.DISP.LOCATION<-as.character(mydaF$myda.DISP.LOCATION)
#Lewis's i-Return Receiving File  -------------1

#First match all SNs with Lewis's iReturn file to have Warranty Status
setwd("C:/Users/Weilun_Chiu/Documents/3PAR/SN")
wb<-loadWorkbook("SN.xlsx")
SN<-readWorksheet(wb, sheet="SN", header=FALSE)
len<-dim(SN)[1]

Result<-data.frame(seq(1, len,1))

for(i in 1:len){
  num<-which(SN[i,1]==mydaF$myda.PART.SERNO)
  if(length(num)>0){
    Result[i,1]<-mydaF[num, 2]
  }else{
    Result[i,1]<-"NA"
  }
}

#Check the NAs with converted SN, and link it with Lewis's iReturn file
num<-which(Result=="NA")
len1<-length(num)

for(i in 1:len1){
  PSN<-substr(SN[num[i],1],6,14)
  LOC<-grep(PSN, mydaF$myda.PART.SERNO)
  
  if(length(LOC)>0){
  Result[num[i],1]<-mydaF$myda.DISP.LOCATION[LOC]
  }else{
  Result[num[i],1]<-"NA"
  }
}

#3PAR
a<-grep("PCMBU", SN$Col1)
b<-grep("PDHWN", SN$Col1)
c<-grep("PDSET", SN$Col1)
d<-grep("PFLKQ", SN$Col1)
POSI<-sort(c(a,b,c,d))
len2<-length(POSI)

setwd("C:/Users/Weilun_Chiu/Documents/3PAR/Assy")
file<-list.files()
Assy<-read.csv(file)
Assy$Assy.<-as.character(Assy$Assy.)
Assy$SPS.<-as.character(Assy$SPS.)

for(i in 1:len2){
  num<-which(SN[POSI[i],1]==Assy$Assy.)
  NodeSN<-Assy[num,2]
  num2<-which(NodeSN==mydaF$myda.PART.SERNO)
  if(length(num2)>0){
    Result[POSI[i],1]<-mydaF$myda.DISP.LOCATION[num2]
  }else{
    Result[POSI[i],1]<-"NA"
  }
}

#Replace RT23 with "IN" and RT24 with "OUT"
len3<-dim(Result)[1]
for(i in 1:len3){
  if(Result[i,1]=="RT23"){
    Result[i,1]<-"IN"
  }else if(Result[i,1]=="RT24"){
    Result[i,1]<-"OUT"
  }else{
      }
}

#Save the workbook
setwd("C:/Users/Weilun_Chiu/Documents/3PAR/SN")
createName(wb, name="Result", formula="SN!$B$1")

writeNamedRegion(wb, Result, name="Result", header=FALSE)

saveWorkbook(wb)
