library(openxlsx)
Data_Path_plot<-file.choose()
Worktable_Gate1 <- read.xlsx(xlsxFile=Data_Path_plot, sheet="Gate1")
Worktable_Gate3 <- read.xlsx(xlsxFile=Data_Path_plot, sheet="Gate3")
Worktable_Gate5 <- read.xlsx(xlsxFile=Data_Path_plot, sheet="Gate5")
dev.new()
Index_col<-length(Worktable_Gate1[3,])-1
for(i in 1:Index_col){
  Gate1_Distance<-c()
  Gate1_time<-c()
  Point_Num<-(length(na.omit(Worktable_Gate1[,i+1]))-10)/5
  for(j in 1:Point_Num){
    Gate1_time[j]<-as.numeric(Worktable_Gate1[j*5+3,i+1])
    Gate1_Distance[j]<-as.numeric(Worktable_Gate1[j*5+7,i+1])
  }
  sp1=spline(Gate1_time,Gate1_Distance,n=1000)
  plot(sp1,xlim=c(0,720),ylim=c(0,80),type="l",xaxt="n")
  x_achse=seq(from=0,to=720,by=20)
  axis(1,x_achse)
  par(new=TRUE)
}
dev.new()
Index_col<-length(Worktable_Gate3[3,])-1
for(i in 1:Index_col){
  Gate3_Distance<-c()
  Gate3_time<-c()
  Point_Num<-(length(na.omit(Worktable_Gate3[,i+1]))-10)/5
  for(j in 1:Point_Num){
    Gate3_time[j]<-as.numeric(Worktable_Gate3[j*5+3,i+1])
    Gate3_Distance[j]<-as.numeric(Worktable_Gate3[j*5+7,i+1])
  }
  sp3=spline(Gate3_time,Gate3_Distance,n=1000)
  plot(sp3,xlim=c(0,720),ylim=c(0,80),type="l",xaxt="n")
  x_achse=seq(from=0,to=720,by=20)
  axis(1,x_achse)
  par(new=TRUE)
}
dev.new()
Index_col<-length(Worktable_Gate5[3,])-1
for(i in 1:Index_col){
  Gate5_Distance<-c()
  Gate5_time<-c()
  Point_Num<-(length(na.omit(Worktable_Gate5[,i+1]))-10)/5
  for(j in 1:Point_Num){
    Gate5_time[j]<-as.numeric(Worktable_Gate5[j*5+3,i+1])
    Gate5_Distance[j]<-as.numeric(Worktable_Gate5[j*5+7,i+1])
  }
  sp5=spline(Gate5_time,Gate5_Distance,n=1000)
  plot(sp5,xlim=c(0,720),ylim=c(0,80),type="l",xaxt="n")
  x_achse=seq(from=0,to=720,by=20)
  axis(1,x_achse)
  par(new=TRUE)
}

