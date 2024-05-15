library(readxl)
library(xts)
library(tidyr)
library(dplyr)
library(zoo)
library(lubridate)
library(DtD)
library(ggplot2)


#removing all previous variables 
rm(list=ls())

#setting WD to code folder  
setwd("C:/Users/veron/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#load functions 
source("Thesis functions.R")

#######################################
#load data
#######################################

#setting WD to data folder
setwd("C:/Users/veron/OneDrive - CBS - Copenhagen Business School/Thesis/3. Data (2010)/1. Extracted Data")
list.files()

#load
SP<-read_excel("S&P.xlsx")
Moody<-read_excel("Moodys.xlsx")

#name first column as Date
names(SP)[1]<-"Date"
names(Moody)[1]<-"Date"

#load the fake firm ID's
Overview<-read_excel("Overview.xlsx")

#create vector with all firms in the main excel file and match with firms in overview
col_namesSP<-names(SP)[-1]
matched_firmsSP<-match(col_namesSP,Overview$Name)

col_namesMoodys<-names(Moody)[-1]
matched_firmsMoodys<-match(col_namesMoodys,Overview$Name)

#create list of unmatched firms
unmatched_firmsSP<-col_namesSP[is.na(matched_firmsSP)]
unmatched_firmsMoodys<-col_namesMoodys[is.na(matched_firmsMoodys)]

#prepare for changing from firm names to fake firm ID's 
new_namesSP<-ifelse(is.na(matched_firmsSP),col_namesSP,Overview$Firm[matched_firmsSP])
new_namesMoodys<-ifelse(is.na(matched_firmsMoodys),col_namesMoodys,Overview$Firm[matched_firmsMoodys])

#rename columns 
names(SP)[-1]<-new_namesSP[!is.na(matched_firmsSP)]
names(Moody)[-1]<-new_namesMoodys[!is.na(matched_firmsMoodys)]

#convert column names to character to make the reshaping work
SP[,-1] <- lapply(SP[,-1], as.character)
Moody[,-1] <- lapply(Moody[,-1], as.character)


#reshape data and arrange by Date and Firm
SP<-SP%>%
  pivot_longer(
    cols=-Date,
    names_to = "Firm",
    values_to = "Rating",
  ) %>%
  arrange(as.numeric(Firm))

#reshape data and arrange by Date and Firm
Moody<-Moody%>%
  pivot_longer(
    cols=-Date,
    names_to = "Firm",
    values_to = "Rating",
  ) %>%
  arrange(as.numeric(Firm))

#replace NA and NR (not rated) values, remove *- and *+, remove NA values
SP<-SP%>%
  mutate(Rating=ifelse(Rating=="#N/A N/A",NA,Rating))%>%
  mutate(Rating=ifelse(Rating=="NR",NA,Rating))%>%
  mutate(Rating = gsub("\\s*\\*[-+]*|[-+]|\\*|\\s+", "", Rating))%>%
  filter(!is.na(Rating))


#replace NA and NR (not rated) values, remove *- and *+, remove NA values
Moody<-Moody%>%
  mutate(Rating=ifelse(Rating=="#N/A N/A",NA,Rating))%>%
  mutate(Rating=ifelse(Rating=="WR",NA,Rating))%>%
  mutate(Rating = gsub("\\s*\\*[-+]*|[-+]|\\*|\\s+", "", Rating))%>%
  filter(!is.na(Rating))

check<-anti_join(Moody,SP,by=c("Date","Firm"))

#define IG and HY 
IG<-c("AAA","AA","A","BBB")
HY<-c("BB","B","CCC","CC","C")

#Define the rating scale
rating_levels <- c("AAA","AA","A","BBB", "BB", "B", "CCC", "CC", "C")

#make dummy indicating whether it is HY or IG rated
SP<-SP%>%
  mutate(ratingdummy=case_when(
    Rating %in% IG ~ 0,
    Rating %in% HY ~ 1,
    TRUE ~ NA_integer_
  ))%>%
  filter(!is.na(Rating))%>%
  mutate(Rating=factor(Rating,levels=rating_levels))



ratingcounts<-table(SP$Rating)
print(ratingcounts)

SP_clean<-SP%>%
  select(Date,Firm,ratingdummy)%>%
  mutate(Firm=as.numeric(Firm))

SP_all<-SP%>%
  select(Date,Firm,ratingdummy,Rating)

#setting WD to code folder  
setwd("C:/Users/veron/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

save(SP_clean,file="SP_Rating.Rdata")

save(SP_all,file="SP_all.Rdata")
