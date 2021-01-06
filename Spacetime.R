library(openxlsx)
FileName <- list.files(path="D:/Diplomarbeit_Sci/Diplomarbeit_Videos/Datensamplezusammen", pattern=".xlsx")
Datapath <- paste("D:/Diplomarbeit_Sci/Diplomarbeit_Videos/Datensamplezusammen",FileName,sep="/")
Vehicle_Gate1<-data.frame()
Vehicle_Gate3<-data.frame()
Vehicle_Gate5<-data.frame()
Index_0<-length(Datapath)
for (i in 1:Index_0){
  Worktable <- read.xlsx(xlsxFile=Datapath[i], sheet="Sheet1")
  col_name<-colnames(Worktable)
  Index<-(nrow(Worktable)-7)/5
  for (k in 1:Index){
    Worktable[k*5+7,1]<-"Distance[m]"
  }
  Index_1<-length(Worktable[1,])-1
  for (j in 1:Index_1){
    Index_2<-(length(na.omit(Worktable[,j+1]))-9)/5
    for (l in 1:Index_2){
      if(l==1){
        Worktable[l*5+7,j+1]<-0
      }else{
        Worktable[l*5+7,j+1]<-as.numeric(Worktable[(l-1)*5+7,j+1])+0.04*as.numeric(Worktable[(l-1)*5+6,j+1])
      }
    }
    Gatenummer<-strsplit(Worktable[2,j+1],"--")
    Gatenummer<-unlist(Gatenummer)
    if(Gatenummer[1]=="1"){
      Col_num1<-ncol(Vehicle_Gate1)
      if(Col_num1<=1){
        Vehicle_Gate1<-data.frame(Worktable[,1])
        colname_Gate1<-c(col_name[1])
        Vehicle_Gate1<-cbind(Vehicle_Gate1,Worktable[,j+1])
        colname_Gate1[2]<-c(col_name[j+1])
      }else{
      Vehicle_Gate1<-cbind(Vehicle_Gate1,Worktable[,j+1])
      colname_Gate1[Col_num1+1]<-c(col_name[j+1])
      }
    }
    if(Gatenummer[1]=="3"){
      Col_num3<-ncol(Vehicle_Gate3)
      if(Col_num3<=1){
        Vehicle_Gate3<-data.frame(Worktable[,1])
        colname_Gate3<-c(col_name[1])
        Vehicle_Gate3<-cbind(Vehicle_Gate3,Worktable[,j+1])
        colname_Gate3[2]<-c(col_name[j+1])
      }else{
      Vehicle_Gate3<-cbind(Vehicle_Gate3,Worktable[,j+1])
      colname_Gate3[Col_num3+1]<-c(col_name[j+1])
      }
    }
    if(Gatenummer[1]=="5"){
      Col_num5=ncol(Vehicle_Gate5)
      if(Col_num5<=1){
        Vehicle_Gate5<-data.frame(Worktable[,1])
        colname_Gate5<-c(col_name[1])
        Vehicle_Gate5<-cbind(Vehicle_Gate5,Worktable[,j+1])
        colname_Gate5[2]<-c(col_name[j+1])
      }else{
      Vehicle_Gate5<-cbind(Vehicle_Gate5,Worktable[,j+1])
      colname_Gate5[Col_num5+1]<-c(col_name[j+1])
      }
    }
  }
  names(Vehicle_Gate1)<-colname_Gate1
  names(Vehicle_Gate3)<-colname_Gate3
  names(Vehicle_Gate5)<-colname_Gate5
  wb = loadWorkbook(Datapath[i])
  sheet1<-addWorksheet(wb, sheetName = "Gate1")
  writeData(wb, sheet=sheet1,x=Vehicle_Gate1,startCol=1,rowNames=FALSE,colNames=TRUE)
  sheet2<-addWorksheet(wb, sheetName = "Gate3")   
  writeData(wb, sheet=sheet2,x=Vehicle_Gate3,startCol=1,rowNames=FALSE,colNames=TRUE)
  sheet3<-addWorksheet(wb, sheetName = "Gate5")  
  writeData(wb, sheet=sheet3,x=Vehicle_Gate5,startCol=1,rowNames=FALSE,colNames=TRUE)
  saveWorkbook(wb, Datapath[i],overwrite = TRUE)
  Vehicle_Gate1<-data.frame()
  Vehicle_Gate3<-data.frame()
  Vehicle_Gate5<-data.frame()
  colname_Gate1<-c()
  colname_Gate3<-c()
  colname_Gate5<-c()
  rm(wb)  
}

