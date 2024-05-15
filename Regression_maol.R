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


#removing all previous variables 
rm(list=ls())

#setting WD to code folder  
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/3. R code")
list.files()

#load functions 
source("Thesis functions.R")

#Load DtD
load("DtD_results.Rdata")
  
#setting WD to data folder
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/2. Data/1. Extracted Data")
list.files()

#######################################
#load data
#######################################

Sector<-read_excel("Sector.xlsx")

#Indepedent variables
Scp1load<-load_thesis_data("Scope 1.xlsx","Scope1","%m/%d/%Y","Overview.xlsx")
Scp1<-Scp1load$df

Scp2load<-load_thesis_data("Scope 2.xlsx","Scope2","%m/%d/%Y","Overview.xlsx")
Scp2<-Scp2load$df

Scp3load<-load_thesis_data("Scope 3.xlsx","Scope3","%m/%d/%Y","Overview.xlsx")
Scp3<-Scp3load$df

#Control variables
Assetload<-load_thesis_data("Total assets EUR.xlsx","Assets","%m/%d/%Y","Overview.xlsx")
Asset<-Assetload$df

Revload<-load_thesis_data("Revenue EUR.xlsx","Revenue","%m/%d/%Y","Overview.xlsx")
Rev<-Revload$df

Liabload<-load_thesis_data("Total liabilities yearly EUR.xlsx","Liabilities","%m/%d/%Y","Overview.xlsx")
Liab<-Liabload$df

prcload<-load_thesis_data("Stock prices (adjusted).xlsx","Price","%m/%d/%Y","Overview.xlsx")
prc<-prcload$df

prcCapload<-load_thesis_data("Stock prices (mktcap).xlsx","Price","%m/%d/%Y","Overview.xlsx")
prcCap<-prcCapload$df

shroutload<-load_thesis_data("Shares outstanding.xlsx","Shrout","%m/%d/%Y","Overview.xlsx")
shrout<-shroutload$df

NProfload<-load_thesis_data("Net profit EUR.xlsx","NetProf","%m/%d/%Y","Overview.xlsx")
NProf<-NProfload$df

BVeqload<-load_thesis_data("BV equity EUR.xlsx","BVEquity","%m/%d/%Y","Overview.xlsx")
BVeq<-BVeqload$df

rf<-get_macro("Euribor 12 m.xlsx","Euribor",2010,2022,"%m/%d/%Y")

MktPrc<-get_macro("Stoxx prices.xlsx","Prices",2010,2022,"%m/%d/%Y")

Infl<-get_macro("Inflation.xlsx","Inflation",2010,2022,"%m/%d/%Y")




#drop load dataframes
rm(list = ls()[grepl("load", ls())])

#setting WD to code folder  
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/3. R code")
list.files()

#########################################
#Get initial firm list 
#########################################
firm_list<-initial_firms(prc,"Price")

#########################################
#sort data
#########################################

#remove data prior to 2010 after 2022, firms not included in the firm list and NA values

Scp1<-filter_firms_and_date(Scp1,2010,2022,firm_list)

Scp2<-filter_firms_and_date(Scp2,2010,2022,firm_list)

Scp3<-filter_firms_and_date(Scp3,2010,2022,firm_list)

Asset<-filter_firms_and_date(Asset,2010,2022,firm_list)

Rev<-filter_firms_and_date(Rev,2010,2022,firm_list)

Liab<-filter_firms_and_date(Liab,2010,2022,firm_list)

prc<-filter_firms_and_date(prc,2010,2022,firm_list)

prcCap<-filter_firms_and_date(prc,2010,2022,firm_list)

shrout<-filter_firms_and_date(shrout,2010,2022,firm_list)

NProf<-filter_firms_and_date(NProf,2010,2022,firm_list)

BVeq<-filter_firms_and_date(BVeq,2010,2022,firm_list)


############################################
#ensure values are defines as numerical data
############################################

Scp1[[3]]<-as.numeric(Scp1[[3]])
Scp2[[3]]<-as.numeric(Scp2[[3]])
Scp3[[3]]<-as.numeric(Scp3[[3]])
Asset[[3]]<-as.numeric(Asset[[3]])
Rev[[3]]<-as.numeric(Rev[[3]])
Liab[[3]]<-as.numeric(Liab[[3]])
prc[[3]]<-as.numeric(prc[[3]])
prcCap[[3]]<-as.numeric(prc[[3]])
shrout[[3]]<-as.numeric(shrout[[3]])
NProf[[3]]<-as.numeric(NProf[[3]])
BVeq[[3]]<-as.numeric(BVeq[[3]])
rf[[2]]<-as.numeric(rf[[2]])
MktPrc[[2]]<-as.numeric(MktPrc[[2]])
Infl[[2]]<-as.numeric(Infl[[2]])

###########################
#Get Market cap
###########################
#merge price and shares outstanding by date and firm 
MktCap <- merge(prcCap, shrout, by = c("Date", "Firm")) %>%
  # sort merged data by firm and then date
  arrange("Firm", "Date") %>%
  # calculate market cap
  mutate(MktCap = Price * Shrout)%>%
  # drop unnecessary variables 
  select(Date,Firm, MktCap, Shrout)

#defining the dates as dates in R 
MktCap$Date<-as.Date(MktCap$Date)

#make a dataframe containing the date of the last observation of the year and the corresponding initial asset volatility for each firm in the data 
MktCapY<-MktCap%>%
  #making a column defining the observation year
  mutate(Year = year(ymd(Date))) %>%
  #grouping data by year and then by firm
  group_by(Year,Firm)%>%
  #takes the max date within each grouping i.e., last trading day of the year
  summarize(Date=max(Date))%>%
  #ad the corresponding initial asset volatility to the last day data
  left_join(MktCap%>%select(Firm,Date,MktCap, Shrout),by=c("Firm","Date"))%>%
  #arrange by Firm and then Date
  arrange("Firm","Date")%>%
  #drop observations with missing initial asset volatility values 
  na.omit(last.day$MktCap)
  
###########################
#Get D/E, ROA, P/B
###########################

#merge Total liabilities, Market Cap, Net profit, Total assets, and Book value equity by date and firm 
Merged_data <- Asset%>%
  full_join(Liab, by=c("Firm","Date"))%>%
  full_join(MktCapY, by=c("Firm","Date"))%>%
  full_join(NProf, by=c("Firm","Date"))%>%
  full_join(BVeq, by=c("Firm","Date"))%>%
  # sort merged data by firm and then date
  arrange("Firm", "Date")%>%
  # calculate market cap
  mutate(DE = Liabilities / MktCap)%>%
  mutate(ROA = NetProf / Assets)%>%
  mutate(PB = MktCap / (BVEquity*Shrout))

DE<-Merged_data%>%
    select(Date, Firm, DE)

ROA<-Merged_data%>%
  select(Date, Firm, ROA)

PB<-Merged_data%>%
  select(Date, Firm, PB)

rm(Merged_data)

###########################
#Return & volatility
###########################
##Stock
prc <- prc %>%
  # Arranging by Date and Firm
  arrange("Date", "Firm") %>%
  # Group by Firm so following calculations are done by firm
  group_by(Firm) %>%
  # Determine daily log return for each firm
  mutate(re = log(Price / dplyr::lag(Price,n=1)))%>%
  # Determine the equity volatility based on a 252-days rolling window (UPDATE if more than 1 year)
  mutate(ve = rollapply(re, width = 252, FUN = sd, fill = NA, align = "right")) %>%
  # Determine the annualized equity volatility
  mutate(ve.annual = ve * sqrt(252)) %>%
  # Ungroup since firm-specific computations are now done
  ungroup() %>%
  # Sort data by firm and then by date
  arrange("Firm", "Date")

##Stoxx
  MktPrc <- MktPrc %>%
    # Arranging by Date
    arrange("Date") %>%
    # Determine daily log return for each firm
    mutate(rMkt = log(Prices / dplyr::lag(Prices,n=1)))%>%
    # Determine the equity volatility based on a 252-days rolling window (UPDATE if more than 1 year)
    mutate(vMkt = rollapply(rMkt, width = 252, FUN = sd, fill = NA, align = "right")) %>%
    # Determine the annualized equity volatility
    mutate(vMkt.annual = vMkt * sqrt(252)) %>%
    # Ungroup since firm-specific computations are now done
    ungroup() %>%
    # Sort data by date
    arrange("Date")

####################################
#Return & volatility to yearly data
####################################

#Stock yearly
prcY<-prc%>%
  #making a column defining the observation year
  mutate(Year = year(ymd(Date))) %>%
  #grouping data by year and then by firm
  group_by(Year,Firm)%>%
  #takes the max date within each grouping i.e., last trading day of the year
  summarize(Date=max(Date))%>%
  #add the corresponding initial asset volatility to the last day data
  left_join(prc%>%select(Firm,Date, Price, re, ve.annual),by=c("Firm","Date"))%>%
  #arrange by Firm and then Date
  arrange("Firm","Date")%>%
  #drop observations with missing initial asset volatility values 
  na.omit(last.day$ve.annual)
  
re<-prcY%>%
    select(Date, Firm, re)

ve<-prcY%>%
  select(Date, Firm, ve.annual)

names(ve)[names(ve) == "ve.annual"] <- "ve"

#Stoxx yearly
MktPrcY <- MktPrc%>%
  #making a column defining the observation year
  mutate(Year = year(ymd(Date))) %>%
  #grouping data by year
  group_by(Year)%>%
  #takes the max date within each grouping i.e., last trading day of the year
  summarize(Date=max(Date))%>%
  #add the corresponding initial asset volatility to the last day data
  left_join(MktPrc%>%select(Date, Prices, rMkt, vMkt.annual),by=c("Date"))%>%
  #arrange by then Date
  arrange("Date")%>%
  #drop observations with missing initial asset volatility values 
  na.omit(last.day$vMkt.annual)

rMkt<-MktPrcY%>%
  select(Date, rMkt)

vMkt<-MktPrcY%>%
  select(Date, vMkt.annual)

names(vMkt)[names(vMkt) == "vMkt.annual"] <- "vMkt"

####################################
#Scope intensity and scope 1+2
####################################

#Calculate intensity and combined 1 and 2 level
ScpItensity <- Scp1%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp3, by=c("Firm","Date"))%>%
  full_join(Rev, by=c("Firm","Date"))%>%
  mutate(Scope1_2=Scope1 + Revenue)%>%
  mutate(Scp1Int=Scope1 / Revenue)%>%
  mutate(Scp2Int=Scope2 / Revenue)%>%
  mutate(Scp3Int=Scope3 / Revenue)%>%
  mutate(Scp1_2Int=Scope1_2 / Revenue)
  
#Extract intensity and scope 1 + 2 level
Scope1_2<-ScpItensity%>%
  select(Date, Firm, Scope1_2)

Scp1Int<-ScpItensity%>%
  select(Date, Firm, Scp1Int)

Scp2Int<-ScpItensity%>%
  select(Date, Firm, Scp2Int)

Scp3Int<-ScpItensity%>%
  select(Date, Firm, Scp3Int)

Scp1_2Int<-ScpItensity%>%
  select(Date, Firm, Scp1_2Int)

#######################################
#Combine data for dataset to regression
#######################################

#Data including scope 3
Reg_dataInclscp3 <- DtD.results%>%
  full_join(Scp1, by=c("Firm","Date"))%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp3, by=c("Firm","Date"))%>%
  full_join(Scope1_2, by=c("Firm","Date"))%>%
  full_join(Scp1Int, by=c("Firm","Date"))%>%
  full_join(Scp2Int, by=c("Firm","Date"))%>%
  full_join(Scp3Int, by=c("Firm","Date"))%>%
  full_join(Scp1_2Int, by=c("Firm","Date"))%>%
  full_join(Asset, by=c("Firm","Date"))%>%
  full_join(Rev, by=c("Firm","Date"))%>%
  full_join(DE, by=c("Firm","Date"))%>%
  full_join(ROA, by=c("Firm","Date"))%>%
  full_join(re, by=c("Firm","Date"))%>%
  full_join(ve, by=c("Firm","Date"))%>%
  full_join(PB, by=c("Firm","Date"))%>%
  full_join(rMkt, by=c("Date"))%>%
  full_join(vMkt, by=c("Date"))%>%
  full_join(Infl, by=c("Date"))%>%
  full_join(rf, by=c("Date"))%>%
  full_join(Sector, by=c("Firm"))%>%
  # sort merged data by firm and then date
  arrange("Date", "Firm")%>%
  # Clean up
  select(-mu, -assetvol)

#Data excluding scope 3
Reg_dataexclscp3<-Reg_dataInclscp3%>%
  select(-Scope3, -Scp3Int)

#Remove rows with NA
Reg_dataInclscp3 <- Reg_dataInclscp3%>%
  na.omit()

Reg_dataexclscp3 <- Reg_dataexclscp3%>%
  na.omit()

#Create dataset with only containing firms with 10 or more observations
FilteredInclscp3 <- Reg_dataexclscp3%>%
  group_by(Firm) %>%
  filter(n() > 10) %>%
  ungroup()

FilteredExclscp3 <- Reg_dataexclscp3 %>%
  group_by(Firm) %>%
  filter(n() > 10)%>%
  ungroup()

###############
#Correlation
###############

##Correlation matrix for scope levels
Corr_scp1<-Reg_dataInclscp3%>%
  select(-DtD,-Date, -Firm, -Year.x, -Year.y, -Sector, -Scp1Int, -Scp2Int, -Scp3Int)

correlation_matrix <- cor(Corr_scp1)

# Generate a table from the correlation matrix
kable(correlation_matrix, caption = "Correlation Matrix Table")

# Export correlation matrix to CSV (activate when needed)
#write.csv(correlation_matrix, "correlation_matrix.csv", row.names = TRUE)

###############
#regressions
###############

#Remove all dateframes expect for reg_data
keep<-c("Reg_dataInclscp3", "Reg_dataexclscp3", "FilteredInclscp3", "FilteredExclscp3")
rm(list = setdiff(ls(), keep))

##Variables
#Date
#Firm
#DtD
#Scope1
#Scope2
#Scope3
#Scope1_2
#Scp1Int
#Scp2Int
#Scp3Int
#Scp1_2Int
#Assets
#Revenue
#DE
#ROA
#re
#ve
#PB
#rMkt
#Inflation
#Euribor
#Sector


Reg_dataexclscp3$Year.x <- as.numeric(Reg_dataexclscp3$Year.x)
Reg_dataexclscp3$Sector_number <- as.numeric(Reg_dataexclscp3$Sector_number)

Reg_dataexclscp3$Firm <- as.factor(Reg_dataexclscp3$Firm)
Reg_dataexclscp3$Year.x <- as.factor(Reg_dataexclscp3$Year.x)
Reg_dataexclscp3$Sector_number <- as.factor(Reg_dataexclscp3$Sector_number)

###Models for levels
#Calculate VIF
modelVif1 <- lm(DtD ~ Scope1 + Scope2 + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + Inflation + Euribor, data = Reg_dataexclscp3)

vif_result1 <- vif(modelVif1)

#Print VIF results
print(vif_result1)

##Model1.1 (Scope 1 and 2 - Levels - Firm fixed)
#Fit a two-way fixed effects model using lm
model1.1 <- lm(DtD ~ Scope1 + Scope2 + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt + Inflation + Euribor + Firm + Year.x - 1, data = Reg_dataexclscp3)

#Adjust for Clustered Standard Errors
vcov1.1 <- multiwayvcov::cluster.vcov(model1.1, 
                                      cbind(Reg_dataexclscp3$Firm, Reg_dataexclscp3$Year.x),
                                      use_white = F, 
                                      df_correction = F)

#Get results
coef_test_results1.1 <- coeftest(model1.1, vcov1.1)

#Extract the results for the first 11 predictors
coef_test1.1 <- coef_test_results1.1[1:13, ]

#View the results
print(coef_test1.1)

##Model1.2 (Scope 1 + 2 - Levels - Firm fixed)
#Fit a two-way fixed effects model using lm
model1.2 <- lm(DtD ~ Scope1_2 + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt + Inflation + Euribor + Firm + Year.x - 1, data = Reg_dataexclscp3)

#Adjust for Clustered Standard Errors
vcov1.2 <- multiwayvcov::cluster.vcov(model1.2, 
                                      cbind(Reg_dataexclscp3$Firm, Reg_dataexclscp3$Year.x),
                                      use_white = F, 
                                      df_correction = F)

#Get results
coef_test_results1.2 <- coeftest(model1.2, vcov1.2)

#Extract the results for the first 11 predictors
coef_test1.2 <- coef_test_results1.2[1:12, ]

#View the results
print(coef_test1.2)

##Model1.3 (Scope 1 and 2 - Levels - Fixed Sector)
#Fit a two-way fixed effects model using lm
model1.3 <- lm(DtD ~ Scope1 + Scope2 + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt+ Inflation + Euribor + Sector_number + Year.x - 1, data = Reg_dataexclscp3)

#Adjust for Clustered Standard Errors
vcov1.3 <- multiwayvcov::cluster.vcov(model1.3, 
                                      cbind(Reg_dataexclscp3$Sector_number, Reg_dataexclscp3$Year.x),
                                      use_white = F, 
                                      df_correction = F)

#Get results
coef_test_results1.3 <- coeftest(model1.3, vcov1.3)

#Extract the results for the first 11 predictors
coef_test1.3 <- coef_test_results1.3[1:13, ]

#View the results
print(coef_test1.3)

##Model1.4 (Scope 1 + 2 - Levels - Fixed Sector)
#Fit a two-way fixed effects model using lm
model1.4 <- lm(DtD ~ Scope1_2 + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt + Inflation + Euribor + Sector_number + Year.x - 1, data = Reg_dataexclscp3)

#Adjust for Clustered Standard Errors
vcov1.4 <- multiwayvcov::cluster.vcov(model1.4, 
                                      cbind(Reg_dataexclscp3$Sector_number, Reg_dataexclscp3$Year.x),
                                      use_white = F, 
                                      df_correction = F)

#Get results
coef_test_results1.4 <- coeftest(model1.4, vcov1.4)

#Extract the results for the first 11 predictors
coef_test1.4 <- coef_test_results1.4[1:12, ]

#View the results
print(coef_test1.4)

###Models for intensity
#Calculate VIF
modelVif2 <- lm(DtD ~ Scp1Int + Scp2Int + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt + Inflation + Euribor, data = Reg_dataexclscp3)

vif_result2 <- vif(modelVif2)

#Print VIF results
print(vif_result2)

##Model2.1 (Scope 1 and 2 - Levels)
#Fit a two-way fixed effects model using lm
model2.1 <- lm(DtD ~ Scp1Int + Scp2Int + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt + Inflation + Euribor + Firm + Year.x - 1, data = Reg_dataexclscp3)

#Adjust for Clustered Standard Errors
vcov2.1 <- multiwayvcov::cluster.vcov(model2.1, 
                                      cbind(Reg_dataexclscp3$Firm, Reg_dataexclscp3$Year.x),
                                      use_white = F, 
                                      df_correction = F)

#Get results
coef_test_results2.1 <- coeftest(model2.1, vcov2.1)

#Extract the results for the first 11 predictors
coef_test2.1 <- coef_test_results2.1[1:13, ]

#View the results
print(coef_test2.1)

##Model2.2 (Scope 1 + 2 - Levels)
#Fit a two-way fixed effects model using lm
model2.2 <- lm(DtD ~ Scp1_2Int + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt + Inflation + Euribor + Firm + Year.x - 1, data = Reg_dataexclscp3)

#Adjust for Clustered Standard Errors
vcov2.2 <- multiwayvcov::cluster.vcov(model2.2, 
                                      cbind(Reg_dataexclscp3$Firm, Reg_dataexclscp3$Year.x),
                                      use_white = F, 
                                      df_correction = F)

#Get results
coef_test_results2.2 <- coeftest(model2.2, vcov2.2)

#Extract the results for the first 11 predictors
coef_test2.2 <- coef_test_results2.2[1:12, ]

#View the results
print(coef_test2.2)

##Model2.3 (Scope 1 and 2 - Levels - Fixed Sector)
#Fit a two-way fixed effects model using lm
model2.3 <- lm(DtD ~ Scp1Int + Scp2Int + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt + Inflation + Euribor + Sector_number + Year.x - 1, data = Reg_dataexclscp3)

#Adjust for Clustered Standard Errors
vcov2.3 <- multiwayvcov::cluster.vcov(model2.3, 
                                      cbind(Reg_dataexclscp3$Sector_number, Reg_dataexclscp3$Year.x),
                                      use_white = F, 
                                      df_correction = F)

#Get results
coef_test_results2.3 <- coeftest(model2.3, vcov2.3)

#Extract the results for the first 11 predictors
coef_test2.3 <- coef_test_results2.3[1:13, ]

#View the results
print(coef_test2.3)

##Model2.4 (Scope 1 + 2 - Levels - Fixed sector)
#Fit a two-way fixed effects model using lm
model2.4 <- lm(DtD ~ Scp1_2Int + Assets + Revenue + DE + ROA + re + ve+ PB+ rMkt + vMkt + Inflation + Euribor + Sector_number + Year.x - 1, data = Reg_dataexclscp3)

#Adjust for Clustered Standard Errors
vcov2.4 <- multiwayvcov::cluster.vcov(model2.4, 
                                      cbind(Reg_dataexclscp3$Sector_number, Reg_dataexclscp3$Year.x),
                                      use_white = F, 
                                      df_correction = F)

#Get results
coef_test_results2.4 <- coeftest(model2.4, vcov2.4)

#Extract the results for the first 11 predictors
coef_test2.4 <- coef_test_results2.4[1:12, ]

#View the results
print(coef_test2.4)

##############################
#Extract coefficients to excel
##############################
# Create a new Excel workbook
wb <- createWorkbook()

# Assuming coef_test1.1, coef_test1.2, coef_test2.1, and coef_test2.2 are your matrices

df_coef_test1_1 <- data.frame(RowLabel = rownames(coef_test1.1), coef_test1.1)
df_coef_test1_2 <- data.frame(RowLabel = rownames(coef_test1.2), coef_test1.2)
df_coef_test1_3 <- data.frame(RowLabel = rownames(coef_test1.3), coef_test1.3)
df_coef_test1_4 <- data.frame(RowLabel = rownames(coef_test1.4), coef_test1.4)
df_coef_test2_1 <- data.frame(RowLabel = rownames(coef_test2.1), coef_test2.1)
df_coef_test2_2 <- data.frame(RowLabel = rownames(coef_test2.2), coef_test2.2)
df_coef_test2_3 <- data.frame(RowLabel = rownames(coef_test2.3), coef_test2.3)
df_coef_test2_4 <- data.frame(RowLabel = rownames(coef_test2.4), coef_test2.4)

# Add each dataframe as a new sheet
addWorksheet(wb, "coef_test1.1")
writeData(wb, "coef_test1.1", df_coef_test1_1)

addWorksheet(wb, "coef_test1.2")
writeData(wb, "coef_test1.2", df_coef_test1_2)

addWorksheet(wb, "coef_test1.3")
writeData(wb, "coef_test1.3", df_coef_test1_3)

addWorksheet(wb, "coef_test1.4")
writeData(wb, "coef_test1.4", df_coef_test1_4)

addWorksheet(wb, "coef_test2.1")
writeData(wb, "coef_test2.1", df_coef_test2_1)

addWorksheet(wb, "coef_test2.2")
writeData(wb, "coef_test2.2", df_coef_test2_2)

addWorksheet(wb, "coef_test2.3")
writeData(wb, "coef_test2.3", df_coef_test2_3)

addWorksheet(wb, "coef_test2.4")
writeData(wb, "coef_test2.4", df_coef_test2_4)

# Save the workbook
saveWorkbook(wb, "Regression_results.xlsx", overwrite = TRUE)

############################
#CAN'T GET THIS TO WORK!!!!#
############################
Reg_dataexclscp3$Year.x <- as.numeric(Reg_dataexclscp3$Year.x)

pdim_info <- pdim(Reg_dataexclscp3, index = c("Firm", "Year.x"))
print(pdim_info)

#model <- plm(DtD ~ Scope1 + ROA + re,
              #data = filtered,
              #model = "within",
              #index = c("Firm", "Year.x"),
              #effect = "twoways") # This specifies both individual and time effects
#summary(model)
#coeftest(model, vcov = vcovHC, type = "HC1")




