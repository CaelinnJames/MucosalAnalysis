#Script written for R 4.1.3 - let me know if you use another R version and run into any issues


### Note: make sure you save data so you have one sample per excel file, with each table on a separate sheet
### Keep the sheet names the same between different files

# Read in data ####
#install.packages("ggplot2") ####Remove the hashtag at the start and run this line if you don't have it already installed.
#install.packages("reshape2") ####Remove the hashtag at the start and run this line if you don't have it already installed.
#install.packages("readxl") ####Remove the hashtag at the start and run this line if you don't have it already installed.
#install.packages("openxlsx") ####Remove the hashtag at the start and run this line if you don't have it already installed.

library(ggplot2)
library(reshape2)
library(openxlsx)

path <- "~/ResultsFolder/"  #### Folder in which your data is stored

read_excel_allsheets <- function(filename) {
  sheets <- readxl::excel_sheets(filename)
  x <-    lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  names(x) <- sheets
  x
}

your_excel_list <- read_excel_allsheets(paste0(path,"Example_Data.xlsx")) #### Change "Example_Data.xslx to name of your file
list2env(your_excel_list, .GlobalEnv)

Sample <- paste(OscAmp[1,1]) ##If you want to manually name a sample, remove everything to the right of the arrow 
                             ##and replace with what you want the sample to be called (in quotation marks)

if(!dir.exists(paste0(path,Sample,"_Graphs"))) {
  dir.create(paste0(path,Sample,"_Graphs"))
}
# Osc amp ####
##  - Caculate tand (G''/G') for Osc amp
OscAmp$tand <- OscAmp$`G''(Pa)`/OscAmp$`G'(Pa)`

your_excel_list[[1]] <- OscAmp
## MAKE SURE YOU DO NOT HAVE THE EXCEL FILE OPEN OR THIS WILL NOT WORK
write.xlsx(your_excel_list, paste0(path,"Example_Data.xlsx"),overwrite = TRUE) #### Change "Example_Data.xslx to name of your file
             


##  - Plot G'/G'' and y*(%) (Osc amp)
OscAmpLong <- melt(OscAmp[,c("G'(Pa)","G''(Pa)")])
OscAmpLong$`y*(%)` <- OscAmp$`y*(%)`
OscAmpLongGraph1_smallaxes <- ggplot(OscAmpLong,aes(y=value,x=`y*(%)`,col=variable))+
  geom_point(size=5)+
  ggtitle(paste0("G' and G'' \n",Sample))+
  ylab("G' and G''")+
  xlab("Shear Strain %")+ 
  theme(legend.title = element_blank())

png(paste0(path,Sample,"_Graphs/Graph1_OscAmpGGYs_ShortAxis_",Sample,".png"),width=700,height=660)
OscAmpLongGraph1_smallaxes
dev.off()

OscAmpLongGraph1_largeaxes <- ggplot(OscAmpLong,aes(y=value,x=`y*(%)`,col=variable))+
  geom_point(size=5)+
  ggtitle(paste0("G' and G'' \n",Sample))+
  ylab("G' and G''")+
  xlab("Shear Strain %")+ 
  theme(legend.title = element_blank())+
  xlim(0,100)+  ##Change the second number to change x axis length
  ylim(0,1500) ##Change the second number to change y axis length

png(paste0(path,Sample,"_Graphs/Graph1_OscAmpGGYs_LargeAxis_",Sample,".png"),width=700,height=660)
OscAmpLongGraph1_largeaxes
dev.off()

##  - G* and y*(%) (Osc amp)
OscAmpLongGraph2_smallaxes <- ggplot(OscAmp,aes(y=`G*(Pa)`,x=`y*(%)`))+
  geom_point(size=5)+
  ggtitle(paste0("G*(Pa) \n",Sample))+
# ylim(0,1500)+     ## Remove the hashtag at the start of this line if you want the y axis to start at 0
  ylab("G*(Pa)")+
  xlab("Shear Strain %")

png(paste0(path,Sample,"_Graphs/Graph2_OscAmpGYs_ShortAxis_",Sample,".png"),width=700,height=660)
OscAmpLongGraph2_smallaxes
dev.off()

OscAmpLongGraph2_largeaxes <- ggplot(OscAmp,aes(y=`G*(Pa)`,x=`y*(%)`))+
  geom_point(size=5)+
  ggtitle(paste0("G*(Pa) \n",Sample))+
  xlim(0,100)+  ##Change the second number to change x axis length
  ylim(0,1500)+ ##Change the second number to change y axis length
  ylab("G*(Pa)")+
  xlab("Shear Strain %")

png(paste0(path,Sample,"_Graphs/Graph2_OscAmpGYs_LargeAxis_",Sample,".png"),width=700,height=660)
OscAmpLongGraph2_largeaxes
dev.off()

##  - δ(°) and y*(%) (Osc amp)
OscAmpLongGraph3_smallaxes <- ggplot(OscAmp,aes(y=`δ(°)`,x=`y*(%)`))+
  geom_point(size=5)+
  ggtitle(paste0("\u03B4(°) \n",Sample))+
  # ylim(0,10)+     ## Remove the hashtag at the start of this line if you want the y axis to start at 0
  ylab("\u03B4(°)")+
  xlab("Shear Strain %")

png(paste0(path,Sample,"_Graphs/Graph3_OscAmpdeltaYs_ShortAxis_",Sample,".png"),width=700,height=660)
OscAmpLongGraph3_smallaxes
dev.off()

OscAmpLongGraph3_largeaxes <- ggplot(OscAmp,aes(y=`δ(°)`,x=`y*(%)`))+
  geom_point(size=5)+
  xlim(0,100)+  ##Change the second number to change x axis length
  ylim(0,100)+ ##Change the second number to change y axis length
  ggtitle(paste0("\u03B4(°) \n",Sample))+
  ylab("\u03B4(°)")+
  xlab("Shear Strain %")
  
png(paste0(path,Sample,"_Graphs/Graph3_OscAmpdeltaYs_LargeAxis_",Sample,".png"),width=700,height=660)
OscAmpLongGraph3_largeaxes
dev.off()

##  - tand and y*(%) (Osc amp)
OscAmpLongGraph4_smallaxes <- ggplot(OscAmp,aes(y=tand,x=`y*(%)`))+
  geom_point(size=5)+
  ggtitle(paste0("tan \u03B4 (G''/G') \n",Sample))+
  # ylim(0,10)+     ## Remove the hashtag at the start of this line if you want the y axis to start at 0
  ylab("tan \u03B4 (G''/G')")+
  xlab("Shear Strain %")

png(paste0(path,Sample,"_Graphs/Graph4_OscAmptandeltaYs_ShortAxis_",Sample,".png"),width=700,height=660)
OscAmpLongGraph4_smallaxes
dev.off()

OscAmpLongGraph4_largeaxes <- ggplot(OscAmp,aes(y=tand,x=`y*(%)`))+
  geom_point(size=5)+
  xlim(0,100)+  ##Change the second number to change x axis length
  ylim(0,100)+ ##Change the second number to change y axis length
  ggtitle(paste0("tan \u03B4 (G''/G') \n",Sample))+
  ylab("tan \u03B4 (G''/G')")+
  xlab("Shear Strain %")

png(paste0(path,Sample,"_Graphs/Graph4_OscAmptandeltaYs_LargeAxis_",Sample,".png"),width=700,height=660)
OscAmpLongGraph4_largeaxes
dev.off()

# OscFreq ####
##  - Plot frequency sweep graph- frequency and δ(°) (Osc freq)
OscFreqGraph_smallaxes <- ggplot(OscFreq,aes(y=`δ(°)`,x=`f(Hz)`))+
  geom_point(size=5)+
  ggtitle(paste0("\u03B4 (°) frequency sweep \n",Sample))+
# ylim(0,14)+     ## Remove the hashtag at the start of this line if you want the y axis to start at 0
  ylab("\u03B4 (°)")+
  xlab("Shear Strain %")

png(paste0(path,Sample,"_Graphs/OscFreq_ShortAxis_",Sample,".png"),width=700,height=660)
OscFreqGraph_smallaxes
dev.off()

OscFreqGraph_largeaxes <- ggplot(OscFreq,aes(y=`δ(°)`,x=`f(Hz)`))+
  geom_point(size=5)+
  ggtitle(paste0("\u03B4 (°) frequency sweep \n",Sample))+
  xlim(0,100)+  ##Change the second number to change x axis length
  ylim(0,1500)+ ##Change the second number to change y axis length
  ylab("\u03B4 (°)")+
  xlab("Shear Strain %")

png(paste0(path,Sample,"_Graphs/OscFreq_LargeAxis_",Sample,".png"),width=700,height=660)
OscFreqGraph_largeaxes
dev.off()

# OscBreakdowns ####
##  - Caculate σ*(Pa) value at 45 degrees (δ(°)) for each of the 4 breakdown oscilations
breakdowns_45 <- data.frame(matrix(ncol=1,nrow=4,dimnames = list(NULL,Sample)),check.names = FALSE)

breakdowns_45[1,1] <- (approx(y = Osc1$`σ*(Pa)`, x = Osc1$`δ(°)`, xout = 45)$y)
breakdowns_45[2,1] <- (approx(y = Osc2$`σ*(Pa)`, x = Osc2$`δ(°)`, xout = 45)$y)
breakdowns_45[3,1] <- (approx(y = Osc3$`σ*(Pa)`, x = Osc3$`δ(°)`, xout = 45)$y)
breakdowns_45[4,1] <- (approx(y = Osc4$`σ*(Pa)`, x = Osc4$`δ(°)`, xout = 45)$y)


if(!file.exists(paste0(path,"AllBreakdowns_45Degrees.txt"))) {
  write.table(breakdowns_45,paste0(path,"AllBreakdowns_45Degrees.txt"),col.names = TRUE,row.names = FALSE,quote = FALSE,sep=",")
} else{
  AllBreakdowns <- read.table(paste0(path,"AllBreakdowns_45Degrees.txt"),header=TRUE,sep=",",check.names = FALSE)
  breakdowns <- cbind(AllBreakdowns,breakdowns_45)
  write.table(breakdowns,paste0(path,"AllBreakdowns_45Degrees.txt"),col.names = TRUE,row.names = FALSE,quote = FALSE,sep=",")
}
### Note: if you run the same sample twice, it will put it in the file twice. 
###If this happens and you don't want it twice, you will have to delete one of the columns, either by loading it in excel 
###or by removing the hashtags in front of the following code and running it **MAKE SURE TO CHANGE THE COLUMN NUMBER TO THE CORRECT COLUMN**
#AllBreakdowns <- read.table(paste0(path,"AllBreakdowns_45Degrees.txt"),header=TRUE)
#AllBreakdowns <- AllBreakdowns[,-c(1)] ### Change 1 to the column number you want to remove - if you want to remove multiple columns, put the numbers inside the curly brackets, and separate the numbers with commas
#write.table(AllBreakdowns,paste0(path,"AllBreakdowns_45Degrees.txt"),col.names = TRUE,row.names = FALSE,quote = FALSE,sep=",")


##  - Plot Osc breakdown intercepts as bar chart
### Note: I would wait until you have run all of your desired samples through the script before doing this part
### Alternatively  you can just make the bar charts on Excel
AllBreakdowns <- read.table(paste0(path,"AllBreakdowns_45Degrees.txt"),header=TRUE,sep=",",check.names = FALSE)

#AllBreakdowns <- AllBreakdowns[,c(1,2,3)] ##Remove the hashtag at the start of this line if you only want to include certain columns in this graph and change the numbers to whichever columns you want to use

AllBreakdowns_Long <- melt(as.matrix(AllBreakdowns))
Breakdowns <- ggplot(data=AllBreakdowns_Long,aes(x=Var1,y=value,fill=Var2))+
  geom_bar(position="dodge",stat="identity")+
  ylab("Sheer stress at breakdown point (Pa)")+
  xlab("Breakdown")+ ## Delete writing to make y axis name blank - LEAVE THE QUOTATION MARKS!
  ggtitle("Breakdown")+
#  scale_fill_manual(values=c("blue","orange","grey"))+ ##R will automatically colour each sample differently - if you want to manually pick the colours then remove the hashtag at the start of this line. You will need to specify the same number of colours as you are using samples - http://www.stat.columbia.edu/~tzheng/files/Rcolor.pdf is a great resource for choosing colours on R
  labs(fill="Sample")

png(paste0(path,"Breakdowns.png"),width=700,height=660) #You will want to fiddle with the file title (make sure to add inside the quotation marks) and the height and width of the plot depending on how many samples you are using
Breakdowns
dev.off()