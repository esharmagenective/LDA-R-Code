#What is the experiment number?
Experiment <- "23028"


#Place Construct List in R working directory
ConstructListFile <- "23028_Constructs.xlsx"


############################################################################################################
User <- as.character(Sys.info()["user"])
User
if (User == "KatieDent") {
  wd <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working"
  RawDataFolder <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working/LDA Raw Data"
  } else {
    if ( User == "EshaSharma") {
      wd <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working"
      RawDataFolder <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working/LDA Raw Data"
    } else {
      if ( User == "LindseyBehrens") {
        wd <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working"
        RawDataFolder <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working/LDA Raw Data"
        Sys.setenv(JAVA_HOME="C:/Program Files/Eclipse Adoptium/jre-11.0.18.10-hotspot")
      }
    }
}




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
library('officer') 
library("readr")
library("pacman")






FileName <- paste(Experiment, "Data",".xlsx", sep= "")





#What Constructs are we looking for?
ConstructList <- read_excel(ConstructListFile, skip = 7)

#subset ConstructList by Control column
notcontrols<-subset(ConstructList,Control!="Y")

search <- notcontrols$Construct
search_list <- as.character(search)
#search_list <- c("TWP280", "TWP10", "285")








#Lists all files in the working folder
List_data <- list.files(path = RawDataFolder,    
                       pattern = "*.xlsx",
                       full.names = TRUE)
List_data #list of all the files in folder

#prints list of files used so we can check to make sure we got all the data.



#ThisFile <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working/LDA Raw Data/230221.xlsx"
#ThisFile <- "C:/Users/LindseyBehrens/OneDrive - Genective/Documents/R Working/LDA Raw Data/230214.xlsx"

#loop through each file to see if anything in the "Agro Strain:" column matches the constructs we are looking for, make a list of those files.
matched_sheets <- data.frame(matrix(ncol=1, nrow=0))
for (ThisFile in List_data) { #loop through each file
  ConList <- read_excel(ThisFile, "Delivery", skip = 3)      
  for (CON in search_list) {
    Result <- str_detect(ConList[,3], CON)
    Result
    if (TRUE %in% Result) {
      matched_sheets <- rbind(matched_sheets, ThisFile)
      print("TRUE")
      break
    } else {
      print("FALSE")
      }
}
}
matched_sheets
write.xlsx(x= matched_sheets, file = FileName, sheetName = "Files Used", append=TRUE)


#Pull all the data and put into one data frame
ExperimentData <- data.frame(matrix(ncol=5, nrow=0))
colnames(ExperimentData) <- c("Date","Pest","Plate","Construct","Rating")
PhytotoxData <- data.frame(matrix(ncol=5, nrow=0))
colnames(PhytotoxData) <- c("Date","Construct","Gene","Day_4","Day_7")
RemoveSheets <- c("Delivery","Phytotox","Expression", "Notes")
for (ThisFile in matched_sheets[,1]) { #loop through each file
  SheetNames <- excel_sheets(ThisFile)
  Sheets <- SheetNames[!(SheetNames %in% RemoveSheets)]
  for (ThisSheet in Sheets) {
    Data <- read_excel(ThisFile, ThisSheet)[,30:34]
    colnames(Data) <- c("Date","Pest","Plate","Construct","Rating")
    ExperimentData <- rbind(ExperimentData, Data)
  }
  if ("Phytotox" %in% SheetNames) {
    ToxData <- read_excel(ThisFile, "Phytotox", skip = 12)[,1:5]
    colnames(ToxData) <- c("Date", "Construct","Gene", "Early_Tox", "Late_Tox")
    ToxData$Date <- as.Date(ToxData$Date)
    ToxData
    PhytotoxData <- rbind(PhytotoxData, ToxData)
  }
}

ExperimentData <-subset(ExperimentData, Plate != "Invalid")
ExperimentData <-subset(ExperimentData, Plate != "invalid")
ExperimentData$Rating <- as.numeric(ExperimentData$Rating)
ExperimentData$Date <- as.Date(ExperimentData$Date)
head(ExperimentData)


Bugs <- as.vector(unique(ExperimentData$Pest)) 
ExperimentData2 <- data.frame(matrix(ncol=5, nrow=0))
colnames(ExperimentData2) <- c("Date","Pest","Plate","Construct","Rating")
for (Bug in Bugs) { #loop through each pest
  BugSet <- subset(ExperimentData, Pest == Bug)
  BugDates <- as.Date(unique(BugSet$Date))
  for (ThisDate in BugDates) {
    DateSet <- subset(BugSet, Date == ThisDate)
    DateCons <- unique(DateSet$Construct)
    Result <- search_list %in% DateCons
    Result
    if (TRUE %in% Result) {
      ExperimentData2 <- rbind(ExperimentData2, DateSet)
    }
  }
}


#prints all data to an Excel File
write.xlsx(x= ExperimentData, file = FileName, sheetName = "ExperimentData", append=TRUE)
write.xlsx(x= PhytotoxData, file = FileName, sheetName = "PhytotoxData", append=TRUE)





