library(extRemes)
library(openxlsx)
FileName <- list.files(path="D:/Diplomarbeit_Sci/Diplomarbeit_Videos/Datensamplezusammen", pattern=".xlsx")
Datapath <- paste("D:/Diplomarbeit_Sci/Diplomarbeit_Videos/Datensamplezusammen",FileName,sep="/")
Rowdaten <- data.frame(Daten_ID=numeric(), Negati_TTC=numeric())
for (i in 1:length(Datapath)){
lll <- read.xlsx(xlsxFile=Datapath[i], sheet="Sheet2")
a <- 100
for (j in 0:(length(lll[,1])-4)/5){                    #get the number of TTC
    if(min(na.omit(as.numeric(lll[j*5+2, ]))) > 0.1){  #ignore the line with negative TTC
    if (min(na.omit(as.numeric(lll[j*5+2, ]))) < a){   #find the minimal TTC in every xlsx-file
       a <- min(na.omit(as.numeric(lll[j*5+2, ])))
    }
    }
}
Rowdaten[i, 1] <- i
Rowdaten[i, 2] <- (-a)
}
fit <- fevd(Negati_TTC, Rowdaten, location.fun= ~1, scale.fun=~1, shape.fun=~1, type=c("GEV"), method=c("Bayesian"), iter=10000, time.units="s")


  
