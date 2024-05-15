#loading packages
library(readxl)
library(openxlsx)
library(xts)
library(tidyr)
library(dplyr)
library(zoo)
library(lubridate)
library(DtD)
library(plm)
library(lmtest)
library(sandwich)
library(car)
library(knitr)
library(car)
library(multiwayvcov)
library(DescTools)

#removing all previous variables 
rm(list=ls())

#setting WD to code folder  
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#load functions 
source("Thesis functions.R")

#setting WD to data folder
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/3. Data (2010)/1. Extracted Data")
list.files()

#######################################
#load data
#######################################

###Firm-specific
prcload<-load_thesis_data("Stock prices (adjusted).xlsx","Price","%m/%d/%Y","Overview.xlsx")
prc<-prcload$df

prcCapload<-load_thesis_data("Stock prices (mktcap).xlsx","Price","%m/%d/%Y","Overview.xlsx")
prcCap<-prcCapload$df

shroutload<-load_thesis_data("Shares outstanding.xlsx","Shrout","%m/%d/%Y","Overview.xlsx")
shrout<-shroutload$df

stdebtload<-load_thesis_data("Short term liabilities quarterly EUR.xlsx","STdebt","%m/%d/%Y","Overview.xlsx")
stdebt<-stdebtload$df

totaldebtload<-load_thesis_data("Total liabilities quarterly EUR.xlsx","Totaldebt","%m/%d/%Y","Overview.xlsx")
totaldebt<-totaldebtload$df

Liabload<-load_thesis_data("Total liabilities yearly EUR.xlsx","Liabilities","%m/%d/%Y","Overview.xlsx")
Liab<-Liabload$df

CurLiabload<-load_thesis_data("Short term liabilities yearly EUR.xlsx","CurLiabilities","%m/%d/%Y","Overview.xlsx")
CurLiab<-CurLiabload$df

Assetload<-load_thesis_data("Total asset.xlsx","Assets","%m/%d/%Y","Overview.xlsx")
Asset<-Assetload$df

CurAssetload<-load_thesis_data("Current Assets.xlsx","CurAssets","%m/%d/%Y","Overview.xlsx")
CurAsset<-CurAssetload$df

Revload<-load_thesis_data("Revenue.xlsx","Revenue","%m/%d/%Y","Overview.xlsx")
Rev<-Revload$df

NProfload<-load_thesis_data("Net profit.xlsx","NetProf","%m/%d/%Y","Overview.xlsx")
NProf<-NProfload$df

RetEarnload<-load_thesis_data("Retained Earnings.xlsx","RetEarning","%m/%d/%Y","Overview.xlsx")
RetEarn<-RetEarnload$df

###Climate
Scp1load<-load_thesis_data("Scope 1.xlsx","Scope1","%m/%d/%Y","Overview.xlsx")
Scp1<-Scp1load$df

Scp2load<-load_thesis_data("Scope 2.xlsx","Scope2","%m/%d/%Y","Overview.xlsx")
Scp2<-Scp2load$df

Scp12load<-load_thesis_data("Scope (E1+2).xlsx","Scope12","%m/%d/%Y","Overview.xlsx")
Scp12<-Scp12load$df

Auditload<-load_thesis_data("Audit.xlsx","Audit","%m/%d/%Y","Overview.xlsx")
Audit<-Auditload$df

###Macro
rf<-get_macro("Euribor 12 m.xlsx","Euribor",2010,2022,"%m/%d/%Y")

MktPrc<-get_macro("Stoxx prices.xlsx","Prices",2010,2022,"%m/%d/%Y")

vMkt<-get_macro("Vstoxx.xlsx","vMkt",2010,2022,"%m/%d/%Y")

Inflation<-get_macro("Inflation.xlsx","Inflation",2010,2022,"%m/%d/%Y")

#Get end date
end_dates<-read_excel("Listing.xlsx")

#Group effects 
Country<-read_excel("Country.xlsx")
Sector<-read_excel("Sector.xlsx")

#setting WD to code folder  
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()
#Misc.
#Load SBTi
load("SBTItarget.Rdata")

#Load Ratings
load("SP_Rating.Rdata")

#######################################
#Combine data
#######################################

Combined<-prc%>%
  full_join(prcCap, by=c("Firm","Date"))%>%
  full_join(shrout, by=c("Firm","Date"))%>%
  full_join(stdebt, by=c("Firm","Date"))%>%
  full_join(totaldebt, by=c("Firm","Date"))%>%
  full_join(Liab, by=c("Firm","Date"))%>%
  full_join(CurLiab, by=c("Firm","Date"))%>%
  full_join(Asset, by=c("Firm","Date"))%>%
  full_join(CurAsset, by=c("Firm","Date"))%>%
  full_join(Rev, by=c("Firm","Date"))%>%
  full_join(NProf, by=c("Firm","Date"))%>%
  full_join(Scp1, by=c("Firm","Date"))%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp12, by=c("Firm","Date"))%>%
  full_join(Audit, by=c("Firm","Date"))%>%
  full_join(rf, by=c("Firm","Date"))%>%
  full_join(MktPrc, by=c("Firm","Date"))%>%
  full_join(vMkt, by=c("Firm","Date"))%>%
  full_join(Inflation, by=c("Firm","Date"))%>%
  full_join(SBTItarget, by=c("Firm","Date"))%>%
  full_join(SP_clean, by=c("Firm","Date"))%>%
  full_join(Country, by=c("Firm","Date"))%>%
  full_join(Sector, by=c("Firm","Date"))
  

