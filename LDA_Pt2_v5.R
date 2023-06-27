###    Alter the names Below    ####
#remember to replace TP with TWP

## ID of Experiment: 
Experiment <- "23059"
Experiment_Title <- "23059: GUN1226A"
Version <- "C"  #change this if you are doing multiple runs



#Place Construct List in R working directory
ConstructListFile <- "23059_Constructs.xlsx"


#Update Controls List if necessary: keep in order of ZsGreen, tox negative control, other controls
ControlConstructs <- c("285", "321", "212", "8", "345", "320", "NT")


#Change x-axis text size
xtext <- 8





####################################################################################################
#End Editing here, just run the script and hope.
####################################################################################################
####################################################################################################
User <- as.character(Sys.info()["user"])
User
if (User == "KatieDent") {
  wd <- "C:/Users/KatieDent/OneDrive - Genective/R working"
  RawDataFolder <- "C:/Users/KatieDent/OneDrive - Genective/R working/LDA Raw Data"
  HiBiT <- read_excel("C:/Users/KatieDent/OneDrive - Genective/Data/Western Blots/HiBiT Results_Read Only.xlsx")
} else {
  if ( User == "EshaSharma") {
    wd <- "C:/Users/EshaSharma/Onedrive - EshaSharma/OneDrive - Genective/Desktop/R Studio/LDA Data"
    RawDataFolder <- "C:/Users/EshaSharma/Onedrive - EshaSharma/OneDrive - Genective/Desktop/R Studio/LDA Data/LDA Raw Data"
    HiBiT <- read_excel("C:/Users/EshaSharma/Onedrive - EshaSharma/OneDrive - Genective/Data/Western Blots/HiBiT Results_Read Only.xlsx")
  } else {
    if ( User == "LindseyBehrens") {
      wd <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working"
      RawDataFolder <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working/LDA Raw Data"
      Sys.setenv(JAVA_HOME="C:/Program Files/Eclipse Adoptium/jre-11.0.19.7-hotspot")
      HiBiT <- read_excel("C:/Users/LindseyBehrens/OneDrive - Genective/Data/Western Blots/HiBiT Results_Read Only.xlsx")
    }
  }
}
setwd(wd)

#Packages to install
library(MASS)
library(car)
library(ggplot2)
library(plyr)
library(gridExtra)
library(ggmosaic)
library(Hmisc)
library(lsr)
library(reshape2)
library(epitools)
library(binom)
library(lme4)
library(data.table) 
library(dplyr)
library(RColorBrewer)
library(cowplot)
library(xlsx)  #needs Java
library(Rmisc)
library(readxl)
library("writexl")
library(multcompView)
library("stringr") 
library('officer') #new 9/22/22
mytheme <-  theme(panel.grid.minor=element_blank(), #gets rid of grey and lines in the middle
                  panel.grid.major=element_blank(), #gets rid of grey and lines in the middle
                  #panel.background=element_rect(fill="white"),#gets rid of grey and lines in the middle
                  panel.border=element_blank(), #gets rid of square going around the entire graph
                  axis.line = element_line(colour = 'black', size = 0.5),#sets the axis line size
                  axis.ticks=element_line(colour = 'black', size = 0.5), #sets the tick lines
                  axis.title.x = element_text(size=16, color="black"), #size of x-axis title
                  axis.title.y = element_text(size=16, color="black"), #size of y-axis title
                  axis.text.x = element_text(size= xtext, color="black", angle = 90, vjust = 0.5, hjust=0.5), #size of x-axis text
                  legend.position="none",
                  strip.text.x = element_text(size = 6, color = "black"),
                  axis.text.y = element_text(size=12, color="black"))#size of y-axis text
Nametheme2 <-  theme(panel.grid.minor=element_blank(), #gets rid of grey and lines in the middle
                    panel.grid.major=element_blank(), #gets rid of grey and lines in the middle
                    #panel.background=element_rect(fill="white"),#gets rid of grey and lines in the middle
                    panel.border=element_blank(), #gets rid of square going around the entire graph
                    axis.line = element_line(colour = 'black', size = 0.5),#sets the axis line size
                    axis.ticks=element_line(colour = 'black', size = 0.5), #sets the tick lines
                    axis.title.x = element_text(size=16, color="black"), #size of x-axis title
                    axis.title.y = element_text(size=16, color="black"), #size of y-axis title
                    axis.text.x = element_text(size= xtext, color="black", angle = 90, vjust = 0.5, hjust=0.5), #size of x-axis text
                    #legend.position="none",
                    axis.text.y = element_text(size=12, color="black"))#size of y-axis text



###########################################################################################################
#Set up Data
###########################################################################################################
#gets today's date and formats into YYMMDD
date <-Sys.Date()
year <- format(date, format="%y")
month <- format(date, format="%m")
day <- format(date, format="%d")
date <-paste(year, month, day, sep="")
date2 <- paste(month, day, year, sep="-")


blackred <- c( "black", "firebrick3")
Colors <- c("limegreen", "firebrick3", "mediumblue", "orange2", "darkorchid3", "#EE1289", "green4", "gold3",  "cyan4", "peru", "lightcoral")
Colors <- rep(Colors, 10)
BlackAndColors <- c("black", Colors)
ShapeSelection<- (c(15, 12, 17, 1, 6, 3, 11))
Shapes <- rep(ShapeSelection, 10)



DataFile <- paste(Experiment, "Data",".xlsx", sep= "")


setwd(wd)
#Set up Feeding data, get rid of NA
rem <- "::"
too <- ":: "
Descriptions <- read_excel(ConstructListFile, "Description", skip = 7)
Descriptions$Construct <- gsub("GV", "", Descriptions$Construct)
Descriptions$Descriptions <- gsub(rem, too, Descriptions$Description)
Descriptions$Descriptions <- gsub("_", "_ ", Descriptions$Descriptions)
Descriptions$Description <- paste(Descriptions$Construct, " ", Descriptions$Descriptions)
Descriptions

FeedingData<- read_excel(DataFile, "ExperimentData")
FeedingData<- FeedingData[,2:6]
FeedingData$Date <- as.factor(FeedingData$Date)
FeedingData$Construct <- as.factor(FeedingData$Construct)
FeedingData$Rating <- as.numeric(FeedingData$Rating)
FeedingData$Construct <- gsub("+285", "", FeedingData$Construct, fixed=T)
FeedingData$Construct <- gsub("PLB", "pLB", FeedingData$Construct, fixed=T)
FeedingData$Construct <- gsub("TP", "TWP", FeedingData$Construct, fixed=T)
FeedingData$Construct <- gsub("TP ", "TWP", FeedingData$Construct, fixed=T)
FeedingData$Construct <- gsub("TWP ", "TWP", FeedingData$Construct, fixed=T)
FeedingData$Construct <- gsub("TWP-", "TWP", FeedingData$Construct, fixed=T)
FeedingData$Rating <- gsub("0", "NA", FeedingData$Rating, fixed=T)
FeedingData$Rating <- as.numeric(FeedingData$Rating)
FeedingData <- na.omit (FeedingData) 
FeedingData <- merge(FeedingData, Descriptions, "Construct")
head(FeedingData)



#Set up Tox data, get rid of NA
ToxData<- read_excel(DataFile, "PhytotoxData")
ToxRows <- nrow(ToxData)
ToxList <- NA
if (ToxRows > 0) {
  ToxData<- ToxData[,2:6]
  ToxData$Date <- as.factor(ToxData$Date)
  ToxData$Construct <- as.factor(ToxData$Construct)
  ToxData$Construct <- gsub("+285", "", ToxData$Construct, fixed=T)
  ToxData$Construct <- gsub("285", "", ToxData$Construct, fixed=T)
  ToxData$Construct <- gsub("PLB", "pLB", ToxData$Construct, fixed=T)
  ToxData$Construct <- gsub("TP", "TWP", ToxData$Construct, fixed=T)
  ToxData$Construct <- gsub("TWP-", "TWP", ToxData$Construct, fixed=T)
  ToxData$Construct <- gsub("TWP ", "TWP", ToxData$Construct, fixed=T)
  ToxData$Construct <- gsub("TP ", "TWP", ToxData$Construct, fixed=T)
  ToxData$Late_Tox <- gsub("0", "NA", ToxData$Late_Tox, fixed=T)
  ToxData$Late_Tox <- as.numeric(ToxData$Late_Tox)
  ToxData<- merge(ToxData, Descriptions, "Construct")
  ToxList <- unique(ToxData$Construct)
  DesList <- unique(Descriptions$Construct)
  ToxList
  DesList
}

AllBugs<- unique(FeedingData$Pest) #list each insect used for feeding assays. needed for loops
ConList <- unique(FeedingData$Construct) #List each construct
TConList <- unique(ToxData$Construct) #List each construct
PList <- unique(FeedingData$POI) #List each protein of interest



#attempt to order data with controls first:
FeedingData_noC <- FeedingData
for (CON in ControlConstructs){
  FeedingData_noC <- subset(FeedingData_noC, Construct != CON)
  }  #This loops through each of the control constructs and keeps everything BUT those in the data. Leaves just experimentals
FeedingData_noC$Construct <- as.factor(FeedingData_noC$Construct)
FeedingData_noC$Construct <- droplevels(FeedingData_noC$Construct) #Gets rid of Constructs with no values so they don't show up on graphs/lists
ExpCon <- as.vector(unique(FeedingData_noC$Construct)) #Making a list of all the experimental constructs
ExpDes <- as.vector(unique(FeedingData_noC$Description)) #Making a list of all the experimental construct descriptions
Bugs <- as.vector(unique(FeedingData_noC$Pest)) 

FeedingData_C <- FeedingData
for (i in 1:length(ExpCon)){
  FeedingData_C <- subset(FeedingData_C, Construct != ExpCon[i])
  }  #This loops through each of the control constructs and keeps just controls
ConCon <- as.vector(unique(FeedingData_C$Construct)) #Making a list of all the control constructs
ConDes <- as.vector(unique(FeedingData_C$Description)) #Making a list of all the control construct descriptions

FeedingData$Description <- factor(FeedingData$Description, levels= c(ConDes,ExpDes)) #uses the lists of controls then the list of experimental constructs to order data.
FeedingData$Construct <- as.factor(FeedingData$Construct)
FeedingData$Construct <- droplevels(FeedingData$Construct) #Drops constructs with no data so they don't show up later





#Use only data for the relevant bugs:
#FeedingData <- subset(FeedingData, Pest = Bugs)
FeedingData <- filter (FeedingData, Pest %in% Bugs)







## Make a folder to print all results, set that as working directory
folder <- paste(Experiment, date, Version, sep="-")
dir.create(folder)
new_wd <- paste(wd, folder, sep= "/")
setwd(new_wd)





###########################################################################################################
#Feeding Data
###########################################################################################################
#1.Feeding Assay: Check significance of 'Date' for each pest with an anova.
#2.make a file with Feeding means+summary stats for each construct in each pest.
#3.Plot for each week but faceted by construct, look for wide variability
FeedPvals <- data.frame(matrix(ncol=4, nrow=0))
for (Bug in Bugs) {
  set <- subset(FeedingData, Pest == Bug) 
  FCons <- unique(set$Construct) 
#1.Anova for week-to-week variation
  for (FC in FCons) {
    set2 <- subset(set, Construct == FC)
    dates <- length(unique(set2$Date))
    if ( dates > 1 ) {
    FpmodelA <- aov(Rating ~ Date, data=set2)
    FanovaA <- summary(FpmodelA)[[1]][["Pr(>F)"]]
    pval <- FanovaA[1] 
    sigfactor <- (pval < 0.05)[1]
    newrow <- c(FC, Bug, pval, sigfactor)
    FeedPvals <- rbind(FeedPvals, newrow)
  }
  }
}
  

for (Bug in Bugs) {
    set <- subset(FeedingData, Pest == Bug)
    dates <- length(unique(set$Date))
#2.feeding averages
  Favg <- set %>%
    group_by(Construct) %>%
    summarize(min = min(Rating),
              q1 = quantile(Rating, 0.25),
              median = median(Rating),
              mean = mean(Rating),
              q3 = quantile(Rating, 0.75),
              'n' = length(`Rating`),
             max = max(Rating))
  write.xlsx(x= Favg, file = "FeedingStats.xlsx", sheetName = Bug, append=TRUE)
}
colnames(FeedPvals) <- c("Construct", "Pest", "PVal", "sig")
FeedingData2 <- merge(FeedingData, FeedPvals, by.x=c("Construct", 'Pest'), by.y=c("Construct", 'Pest'))
head(FeedingData2)


MeanList <- Descriptions
MeanList$Description <- NULL
for (Bug in Bugs) {
  Sheet <- read_excel("FeedingStats.xlsx", Bug)
  set <- Sheet[, c("Construct", "mean")]
  set$mean <- as.numeric(format(round(set$mean, 2), nsmall = 2))
  set[,paste(Bug, "Mean", sep= "_")] <- set$mean
  set$mean <- NULL
  MeanList <- merge(MeanList, set, by = "Construct", all= TRUE)
}
MeanList$mean <- NULL
MeanList$Construct <- as.character(MeanList$Construct)
OrderList <- c(ConCon,ExpCon)
MeanList <- MeanList[match(OrderList, MeanList$Construct),]
write.xlsx(x= MeanList, file = "MeanList.xlsx")


NList <- Descriptions
NList$Description <- NULL
for (Bug in Bugs) {
  Sheet <- read_excel("FeedingStats.xlsx", Bug)
  set <- Sheet[, c("Construct", "n")]
  #set$n <- as.numeric(format(round(set$n, 2), nsmall = 2))
  set[,paste(Bug, "n", sep= "_")] <- set$n
  set$n <- NULL
  NList <- merge(NList, set, by = "Construct", all= TRUE)
}
NList$n <- NULL
NList$Construct <- as.character(NList$Construct)
OrderList <- c(ConCon,ExpCon)
NList <- NList[match(OrderList, NList$Construct),]
write.xlsx(x= NList, file = "n_List.xlsx")



for (Bug in Bugs) {
  set <- subset(FeedingData2, Pest == Bug)
  head(set)
  dates <- length(unique(set$Date))
#3.Plot for each week faceted by construct, show variation
name <- paste (Bug, " Weekly Variation", sep="" )
if ( dates > 1 ) {
  facetPlot <- ggplot(data=set, aes(x=Date, y=Rating, group= Description, colour= sig)) + 
    scale_color_manual(values = blackred) +
    ggtitle(name) +
    geom_point(size=1, position = position_jitter(w=0.3, h=0.05)) + 
    stat_smooth(aes(group=1)) + 
    stat_summary(aes(group=1)) + 
    facet_grid( ~ Description, labeller = labeller(Description = label_wrap_gen(10))) +
    mytheme
  name <- paste(Bug,"facet", ".png", sep="")
  png(name, units="in", width=12, height=6, res= 300)
  plot(facetPlot)
  dev.off()
  }
  facetPlot <- "na"
}



#Dot Plots for each individual insect with standard error bars
for (Bug in Bugs) {
  set <- subset(FeedingData, Pest == Bug) 
  #Make letters to represent stat. sig:
  mod_feed <- lm((Rating) ~ Construct, data = set)
  feed.av <- aov(mod_feed)
  #tukey post hoc test
  TUKEY <- TukeyHSD(feed.av)
  plot(TUKEY) #shows 95% CI of a pairwise contrast. 
  #make the letters for plotting:
  cld <- multcompLetters4(feed.av, TUKEY)
  # table with factors and 3rd quantile
  dt <- group_by(set, Construct) %>%
    summarise(w=max(Rating), sd = sd(Rating)) 
  # extracting the compact letter display and adding to the Tk table
  cld <- as.data.frame.list(cld$Construct)
  cld$Construct <- rownames(cld)
  dt <- merge(dt, cld, by= "Construct")
  dt$POI <- dt$Construct
  dt <- merge(dt, Descriptions, "Construct")
  dt_length <- length(dt$POI)
  list <- vector( "character" , dt_length )
  dt$StatisticalGroup <- list

  
Fplot <- ggplot(data=set, aes(x=Description, y=Rating, color = Description))  +
  scale_color_manual(values = BlackAndColors) +
  ggtitle(Bug) +
  geom_point(data=set, size=1.5, position = position_jitter(w=0.3, h=0.04), aes(shape = Date)) +
  scale_shape_manual(values=Shapes)+
  guides(color = "none") +
  stat_summary(fun.data = mean_cl_normal, geom = "errorbar", position = position_dodge(width = 0.90),width=.2) +
  stat_summary(fun = mean, colour="black") +
  geom_hline(yintercept=2, linetype="dotted", color = "red") +
  geom_text(data=dt,aes(x=Description, y=3.2, label=Letters, colour = "'Statistical Groups'"), show.legend = FALSE) + #y=is where the letter sit on the yaxis. 
  scale_x_discrete(labels = function(Construct) str_wrap(Construct, width = 5)) +
  Nametheme2
Fplot
name <- paste(Bug,"feeding2", ".png", sep="")
png(name, units="in", width=10, height=6, res= 300)
plot(Fplot)
dev.off()
  }
  


#Pairwise wilcox test (pair-wise Mann-Whitney U test) to classify significance
for (Bug in Bugs) {
  set <- subset(FeedingData, Pest == Bug) 
  wil <- pairwise.wilcox.test(set$Rating, set$Construct, p.adj="bonf")
  wildf <- wil$p.value
  write.xlsx(x= wildf, file = "FeedingWilcox.xlsx", sheetName = Bug, append=TRUE)
}




Allplot <- ggplot(data=FeedingData, aes(x=Description, y=Rating,  color = Pest))  +
  scale_color_manual(values = Colors) +
  ggtitle("Feeding Assay") +
  stat_summary(fun.data = mean_cl_normal, geom = "errorbar", position = position_dodge(width = 0.5),width=.4) +
  geom_hline(yintercept=2, linetype="dotted", color = "red") +
  scale_x_discrete(labels = function(Description) str_wrap(Description, width = 5)) +
  Nametheme2
Allplot
png("AllFeedingPlot.png", units="in", width=10, height=6, res= 300)
plot(Allplot)
dev.off()








####################################################################################################
#Phytotox
####################################################################################################
#attempt to order data with controls first:

if (length(ToxList) > 2) {
ToxData_noC <- ToxData
for (i in 1:length(ControlConstructs)){
  ToxData_noC <- subset(ToxData_noC, Construct != ControlConstructs[i])}  #This loops through each of the control constructs and keeps everything BUT those in the data. Leaves just experimentals
ToxData_noC$Construct <- as.factor(ToxData_noC$Construct)
ToxData_noC$Construct <- droplevels(ToxData_noC$Construct) #Gets rid of Constructs with no values so they don't show up on graphs/lists
ToxExpCon <- as.vector(unique(ToxData_noC$Construct)) #Making a list of all the experimental constructs
ToxExpDes <- as.vector(unique(ToxData_noC$Description)) #Making a list of all the experimental construct descriptions

ToxData_C <- ToxData
for (i in 1:length(ToxExpCon)){
  ToxData_C <- subset(ToxData_C, Construct != ToxExpCon[i])}  #This loops through each of the control constructs and keeps just controls
ToxConDes <- as.vector(unique(ToxData_C$Description)) #Making a list of all the control construct descriptions

ToxData$Description <- factor(ToxData$Description, levels= c(ToxConDes,ToxExpDes)) #uses the lists of controls then the list of experimental constructs to order data.
ToxData$Construct <- as.factor(ToxData$Construct)
ToxData$Construct <- droplevels(ToxData$Construct) #Drops constructs with no data so they don't show up later
}
ToxData$Late_Tox <- as.numeric(ToxData$Late_Tox)
ToxData$Early_Tox <- as.numeric(ToxData$Early_Tox)

if (length(ToxList) > 2) {
#Tox averages
Tavg <- ToxData %>%
    group_by(Construct) %>%
    summarize(Late_min = min(Late_Tox),
              Late_q1 = quantile(Late_Tox, 0.25),
              Late_median = median(Late_Tox),
              Late_Tox = mean(Late_Tox),
              Late_q3 = quantile(Late_Tox, 0.75),
              Late_max = max(Late_Tox),
              n = length(Late_Tox),
              Early_min = min(Early_Tox),
              Early_q1 = quantile(Early_Tox, 0.25),
              Early_median = median(Early_Tox),
              Early_Tox = mean(Early_Tox),
              Early_q3 = quantile(Early_Tox, 0.75),
              Early_max = max(Early_Tox))
  write.xlsx(x= Tavg, file = "ToxStats.xlsx")
Tavg2 <- Tavg[,c("Construct", "Early_Tox", "Late_Tox", "n")]

ToxMeanList <- Descriptions
ToxMeanList$Description <- NULL
ToxMeanList <- merge(ToxMeanList, Tavg2, "Construct")
ToxMeanList
write.xlsx(x= ToxMeanList, file = "ToxMeanList.xlsx")
}


if (length(ToxList) > 2) {
#Make letters to represent stat. sig:
  mod_tox <- lm((Late_Tox) ~ Construct, data = ToxData)
  tox.av <- aov(mod_tox)
#tukey post hoc test
  TUKEY2 <- TukeyHSD(tox.av)
  plot(TUKEY2) #shows 95% CI of a pairwise contrast. 
#make the letters for plotting:
  cld2 <- multcompLetters4(tox.av, TUKEY2)
# table with factors and 3rd quantile
  dt2 <- group_by(ToxData, Construct) %>%
    summarise(w=max(Late_Tox), sd = sd(Late_Tox)) 
# extracting the compact letter display and adding to the Tk table
  cld2 <- as.data.frame.list(cld2$Construct)
  cld2$Construct <- rownames(cld2)
  dt2 <- merge(dt2, cld2, by= "Construct")  
  dt2 <- merge(dt2, Descriptions, "Construct")

#Phytotoxicity Dot Plots
Tplot <- ggplot(data=ToxData, aes(x=Description, y=Late_Tox, colour= Description))  +
  ggtitle("Phytotoxicity Day 7") +
  geom_point(data=ToxData, size=2, position = position_jitter(w=0.3, h=0.04), aes(shape = Date)) +
  scale_shape_manual(values=Shapes)+
  guides(color = "none") +
  coord_cartesian(ylim = c(1, 4)) +
  stat_summary(fun.data = mean_cl_normal, geom = "errorbar", position = position_dodge(width = 0.90),width=.2) +
  stat_summary(fun = mean, colour="black") +
  ylab("Phytotoxicity Rating") +
  geom_hline(yintercept=2, linetype="dotted", color = "red") +
  scale_x_discrete(labels = function(Description) str_wrap(Description, width = 5)) +
  Nametheme2
if (exists("dt2") == TRUE){
  Tplot <- Tplot+ 
  geom_text(data=dt2,aes(x=Description, y=4.2, label=Letters, colour = "'Statistical Groups'"), show.legend = FALSE) + #y=is where the letter sit on the yaxis. 
  scale_color_manual(values = BlackAndColors)
} else {
  Tplot <- Tplot+scale_color_manual(values = Colors)
    }
Tplot
png("ToxPlot.png", units="in", width=10, height=6, res= 300)
plot(Tplot)
dev.off()


Twil <- pairwise.wilcox.test(ToxData$Late_Tox, ToxData$Construct, p.adj="bonf")
Twildf <- Twil$p.value
write.xlsx(x= Twildf, file = "ToxWilcox.xlsx")
}



####################################################################################################
#Make a pretty table
####################################################################################################
HiBiT$Construct <- gsub(" ", "", HiBiT$Construct, fixed=T)
HiBiT <- HiBiT %>% select(-Plasmid)
HiBiT <- HiBiT %>% select(-Number)
HiBiT <- HiBiT %>% select(-POI)
HiBiT <- HiBiT %>% select(-ends_with(" Date"))
HiBiT <- HiBiT %>% select(-ends_with(" Expression"))
#ExpCon  #list of all the experimental constructs
HiBiT2 <- subset(HiBiT, Construct %in% ExpCon)
HiBiT2 <- melt(HiBiT2, id= c('Construct', 'Description'))
HiBiT3 <- HiBiT2 %>%
  group_by(Construct) %>%
  summarize(Mean_LUM = mean(value, na.rm=T))
HiBiT3$Mean_Expression <- with(HiBiT3, ifelse(Mean_LUM >20000,"Very High",
                                              ifelse(Mean_LUM >4999,"High", 
                                                     ifelse(Mean_LUM >999,"Medium",
                                                            ifelse(Mean_LUM >99,"Low",
                                                                   ifelse(Mean_LUM >19, "Very Low",
                                                                          ifelse(Mean_LUM >0, "None", "")))))))



#Uing flextable package
#https://ardata-fr.github.io/officeverse/officer-for-powerpoint.html
#https://ardata-fr.github.io/flextable-book/table-design.html
library("flextable")
library(scales)

if (length(ToxList) > 2) {
  MeanMerge <- merge(MeanList, ToxMeanList, c("Construct", "Descriptions", "Control"), all = TRUE )
  MeanMerge <- MeanMerge %>% select(-n)
  MeanMerge <- MeanMerge %>% select(-Control)
  ToxCols <- as.vector(colnames(MeanMerge %>% select(ends_with('_Tox'))))
} else {
  MeanMerge <- MeanList
  MeanMerge <- MeanMerge %>% select(-Control)
}



ColOrder <- c("Construct","Descriptions", "FAW_Mean","ECB_Mean","CEW_Mean", "SL_Mean", "BCW_Mean", "SCR_Mean", "Cry1Fa-rFAW_Mean", "Vip3a-rFAW_Mean", "Early_Tox", "Late_Tox")
ColOrder <- ColOrder[ColOrder %in% colnames(MeanMerge)]
MeanMerge <- MeanMerge[ColOrder]
MergeCols <- as.vector(colnames(MeanMerge %>% select(ends_with('_Mean'))))
MeanMerge <- merge(MeanMerge, HiBiT3, "Construct", all = TRUE )
MeanMerge <- MeanMerge %>% select(-Mean_LUM)



colourer <- col_numeric(
  palette = c("olivedrab3", "gold1", "firebrick1"),
  domain = c(1, 3))
colourer2 <- col_numeric(
  palette = c("olivedrab3", "gold1", "firebrick1"),
  domain = c(1, 4))

ft <- flextable(MeanMerge)
ft <- bg( ft, bg = colourer,
          j = MergeCols, part = "body")
ft <- bg( ft, bg = colourer2,
          j = ToxCols, part = "body")
ft <- ft %>% 
  bg(~ Mean_Expression == "None", bg = "firebrick", j= "Mean_Expression")  %>% 
  bg(~ Mean_Expression == "Very Low", bg = "firebrick1", j= "Mean_Expression") %>% 
  bg(~ Mean_Expression == "Low", bg = "orange", j= "Mean_Expression") %>% 
  bg(~ Mean_Expression == "Medium", bg = "gold1", j= "Mean_Expression") %>% 
  bg(~ Mean_Expression == "High", bg = "olivedrab3", j= "Mean_Expression") %>% 
  bg(~ Mean_Expression == "Very High", bg = "olivedrab", j= "Mean_Expression") %>% 
  bg(~ Mean_Expression == "Detected", bg = "olivedrab3", j= "Mean_Expression")  %>% 
  bg(~ Mean_Expression == "Faint", bg = "orange", j= "Mean_Expression") %>% 
  bg(~ Mean_Expression == "Not Detected", bg = "firebrick", j= "Mean_Expression")
ft <- theme_booktabs(ft, bold_header = TRUE) 
ft <- align(ft, align = "center")
ft <- fontsize(ft, size= 10, part = "all")
ft <- set_table_properties(ft, layout = "fixed")











####################################################################################################
#Make a Power Point Presentation
####################################################################################################
bugpic <- file.path(wd, "BugsPicts.png")
leafpic <- file.path(wd, "LeafPicts.png")
logo <-  file.path(wd, "GenLogo.png")
score <- file.path(wd, "LDAScoring.png")
setwd(new_wd)


#Set location for bug and leaf graphics
loc_bug <-ph_location(left = 0.3, top = 1.9, width = .6, height = 4.4)
loc_leaf <- ph_location(left = 0.1, top = 1.9, width = .6, height = 4.4)
loc_logo <- ph_location(left = 8, top = 7, width = 1.7, height = .5)
loc_graph <- ph_location(left = .48, top = 1.55, width = 9.5, height = 5.28)
loc_table <- ph_location(left = .3, top = 1.55, width = 9.5, height = 4.7)


# Title Slide
slides <- read_pptx()
slides <- add_slide(slides, layout= "Title and Content")
PPtitle <- fpar(ftext(Experiment_Title, fp_text(bold = TRUE, font.size = 50)), fp_p = fp_par(text.align = "center"))
slides <- ph_with(slides, value = PPtitle, location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo, width = 9, height = 2.5), location = ph_location_type(type = "body"), use_loc_size = FALSE)


#Scoring
slides <- add_slide(slides, layout=  "Two Content")
slides <- ph_with(slides, value = "Leaf Disk Assay Overview", location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))
slides <- ph_with(slides, external_img(score), location = ph_location_type(type = "body", id = 2), use_loc_size = TRUE)
LDA_Text <- block_list(
  fpar(ftext("- Leaf disks are scored from 3 (high feeding) to 1 (no feeding)")),
  fpar(ftext("- Score based on % leaf area remaining")),
  fpar(ftext("- We run 8 biological replicates per construct per week for at least 2 weeks")))
slides <- ph_with(slides, value = LDA_Text, location = ph_location_type(type = "body", id = 1))



#Purpose
slides <- add_slide(slides, layout= "Title and Content")
slides <- ph_with(slides, value = "Purpose", location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))


#Background
slides <- add_slide(slides, layout= "Title and Content")
slides <- ph_with(slides, value = "Background", location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))


#Feeding Stat Summary Slide
slides <- add_slide(slides, layout= "Title and Content")
slides <- ph_with(slides, value = "Transient Assay Summary", location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = MeanList, location = loc_table)
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))


#All Feeding Slide
slides <- add_slide(slides, layout= "Title and Content")
slides <- ph_with(slides, value = "All Feeding Data", location = ph_location_type(type = "title"))
slides <- ph_with(slides, external_img("AllFeedingPlot.png"), location = loc_graph)
slides <- ph_with(slides, external_img(bugpic), location = loc_bug)
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))


#Feeding Data Slides
for (Bug in Bugs) {
  slides <- add_slide(slides, layout= "Title and Content")
  facetpict <- paste(Bug,"facet", ".png", sep="")
  if (file.exists(facetpict) == TRUE ) {
  title_facet <- paste (Bug, " Weekly Variation", sep="" )
  slides <- ph_with(slides, value = title_facet, location = ph_location_type(type = "title"))
  slides <- ph_with(slides, external_img(facetpict), location = loc_table)
  slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
  slides <- ph_with(slides, external_img(logo), location = loc_logo)
  slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))
  }
  slides <- add_slide(slides, layout= "Title and Content")
  feedpict <- paste(Bug,"feeding2", ".png", sep="")
  if (file.exists(feedpict) == TRUE ) {
  title_feed <- paste (Bug, " Feeding Assay", sep="" )
  slides <- ph_with(slides, value = title_feed, location = ph_location_type(type = "title"))
  slides <- ph_with(slides, external_img(feedpict), location = loc_graph)
  slides <- ph_with(slides, external_img(bugpic), location = loc_bug)
  slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
  slides <- ph_with(slides, external_img(logo), location = loc_logo)
  slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))
  } else {
    slides <- add_slide(slides, layout= "Title and Content")
    feedpict <- paste(Bug,"feeding", ".png", sep="")
    if (file.exists(feedpict) == TRUE ) {
      title_feed <- paste (Bug, " Feeding Assay", sep="" )
      slides <- ph_with(slides, value = title_feed, location = ph_location_type(type = "title"))
      slides <- ph_with(slides, external_img(feedpict), location = loc_graph)
      slides <- ph_with(slides, external_img(bugpic), location = loc_bug)
      slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
      slides <- ph_with(slides, external_img(logo), location = loc_logo)
      slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))
    }
  }
}


if (length(ToxList) > 2) {
#Phytotox Results Slide
slides <- add_slide(slides, layout= "Title and Content")
slides <- ph_with(slides, value = "Phytotoxicity Data", location = ph_location_type(type = "title"))
slides <- ph_with(slides, external_img("ToxPlot.png"), location = loc_graph)
slides <- ph_with(slides, external_img(leafpic), location = loc_leaf)
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))

#Phytotox Stat Summary Slide
slides <- add_slide(slides, layout= "Title and Content")
slides <- ph_with(slides, value = "Phytotoxicity Results", location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))
slides <- ph_with(slides, value = ToxMeanList, location = loc_table)
}


#Next Steps Slide
slides <- add_slide(slides, layout= "Title and Content")
slides <- ph_with(slides, value = "Next Steps", location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = "Add Next Steps here", location = ph_location_type(type = "body"))
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))


# n= slide
slides <- add_slide(slides, layout= "Title and Content")
slides <- ph_with(slides, value = "Replicates Used", location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = NList, location = loc_table)
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))


#Stats
slides <- add_slide(slides, layout=  "Title and Content")
slides <- ph_with(slides, value = "Statistics Used", location = ph_location_type(type = "title"))
slides <- ph_with(slides, value = "Confidential", location = ph_location_type(type = "ftr"))
slides <- ph_with(slides, external_img(logo), location = loc_logo)
slides <- ph_with(slides, value = date2, location = ph_location_type(type = "dt"))
LDA_Text <- block_list(
  fpar(ftext("- An Anova is used to look for week-to-week variation.", fp_text(font.size = 28))),
  fpar(ftext("- Significant week-to-week variation is colored red.", fp_text(font.size = 28))),
  fpar(ftext("- A Tukey test is used to look for differences between feeding results between constucts for each pest.", fp_text(font.size = 28))),
  fpar(ftext("- The Tukey test generates letters on the feeding graphs to group the constructs.", fp_text(font.size = 28))),
  fpar(ftext("- The feeding graphs also include a black dot for the mean and error bars (stardard error)", fp_text(font.size = 28))))
slides <- ph_with(slides, value = LDA_Text, location = ph_location_type(type = "body"))



ReportTitle <- paste(Experiment, "Report", date2, ".pptx", sep=" ")
print(slides, target = ReportTitle) 


