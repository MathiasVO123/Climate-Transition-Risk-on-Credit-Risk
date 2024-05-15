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
library(xtable)

#removing all previous variables 
rm(list=ls())

#setting WD to code folder  
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#load functions 
source("Thesis functions.R")

#Load DtD1
load("Clean_DtD_neg_rm.Rdata")

#Load DtD2
load("Clean_DtD_neg_zero.Rdata")

#Load DtD3
load("DtD_wip.Rdata")

#Load SBTi
load("SBTItarget.Rdata")

#Load Ratings
load("SP_Rating.Rdata")

#setting WD to data folder
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/3. Data (2010)/1. Extracted Data")
list.files()

#######################################
#load data
#######################################

Country<-read_excel("Country.xlsx")
Sector<-read_excel("Sector.xlsx")
end_dates_data<-read_excel("Listing.xlsx")

#Indepedent variables
Scp1load<-load_thesis_data("Scope 1.xlsx","Scope1","%m/%d/%Y","Overview.xlsx")
Scp1<-Scp1load$df

Scp2load<-load_thesis_data("Scope 2.xlsx","Scope2","%m/%d/%Y","Overview.xlsx")
Scp2<-Scp2load$df

Scp3load<-load_thesis_data("Scope 3.xlsx","Scope3","%m/%d/%Y","Overview.xlsx")
Scp3<-Scp3load$df

Scp12load<-load_thesis_data("Scope (E1+2).xlsx","Scope12","%m/%d/%Y","Overview.xlsx")
Scp12<-Scp12load$df

Auditload<-load_thesis_data("Audit.xlsx","Audit","%m/%d/%Y","Overview.xlsx")
Audit<-Auditload$df
Audit <- Audit %>%
  mutate(Audit = ifelse(!is.na(Audit), 1, 0))

#Control variables
Assetload<-load_thesis_data("Total asset.xlsx","Assets","%m/%d/%Y","Overview.xlsx")
Asset<-Assetload$df

CurAssetload<-load_thesis_data("Current Assets.xlsx","CurAssets","%m/%d/%Y","Overview.xlsx")
CurAsset<-CurAssetload$df

Revload<-load_thesis_data("Revenue.xlsx","Revenue","%m/%d/%Y","Overview.xlsx")
Rev<-Revload$df

Liabload<-load_thesis_data("Total liabilities yearly EUR.xlsx","Liabilities","%m/%d/%Y","Overview.xlsx")
Liab<-Liabload$df

CurLiabload<-load_thesis_data("Current Liabilities.xlsx","CurLiabilities","%m/%d/%Y","Overview.xlsx")
CurLiab<-CurLiabload$df

RetEarnload<-load_thesis_data("Retained Earnings.xlsx","RetEarning","%m/%d/%Y","Overview.xlsx")
RetEarn<-RetEarnload$df

prcload<-load_thesis_data("Stock prices (adjusted).xlsx","Price","%m/%d/%Y","Overview.xlsx")
prc<-prcload$df

prcCapload<-load_thesis_data("Stock prices (mktcap).xlsx","Price","%m/%d/%Y","Overview.xlsx")
prcCap<-prcCapload$df

shroutload<-load_thesis_data("Shares outstanding.xlsx","Shrout","%m/%d/%Y","Overview.xlsx")
shrout<-shroutload$df

NProfload<-load_thesis_data("Net profit.xlsx","NetProf","%m/%d/%Y","Overview.xlsx")
NProf<-NProfload$df

BVeqload<-load_thesis_data("BV equity.xlsx","BVEquity","%m/%d/%Y","Overview.xlsx")
BVeq<-BVeqload$df

rf<-get_macro("Euribor 12 m.xlsx","Euribor",2010,2022,"%m/%d/%Y")

MktPrc<-get_macro("Stoxx prices.xlsx","Prices",2010,2022,"%m/%d/%Y")

vMkt<-get_macro("Vstoxx.xlsx","vMkt",2010,2022,"%m/%d/%Y")

Inflation<-get_macro("Inflation.xlsx","Inflation",2010,2022,"%m/%d/%Y")

#drop load dataframes
rm(list = ls()[grepl("load", ls())])

#setting WD to code folder  
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#########################################
#Get initial firm list 
#########################################

firm_list<-initial_firms(prc,"Price")
save(firm_list,file="firm_list.Rdata")

#########################################
#sort data
#########################################

#remove data prior to 2010 after 2022, firms not included in the firm list and NA values

Scp1<-filter_firms_and_date(Scp1,2009,end_dates_data,firm_list)

Scp2<-filter_firms_and_date(Scp2,2009,end_dates_data,firm_list)

Scp3<-filter_firms_and_date(Scp3,2009,end_dates_data,firm_list)

Scp12<-filter_firms_and_date(Scp12,2009,end_dates_data,firm_list)

Asset<-filter_firms_and_date(Asset,2010,end_dates_data,firm_list)

CurAsset<-filter_firms_and_date(CurAsset,2010,end_dates_data,firm_list)

Rev<-filter_firms_and_date(Rev,2010,end_dates_data,firm_list)

Liab<-filter_firms_and_date(Liab,2010,end_dates_data,firm_list)

CurLiab<-filter_firms_and_date(CurLiab,2010,end_dates_data,firm_list)

RetEarn<-filter_firms_and_date(RetEarn,2010,end_dates_data,firm_list)

prc<-filter_firms_and_date(prc,2010,end_dates_data,firm_list)

prcCap<-filter_firms_and_date(prc,2010,end_dates_data,firm_list)

shrout<-filter_firms_and_date(shrout,2010,end_dates_data,firm_list)

NProf<-filter_firms_and_date(NProf,2010,end_dates_data,firm_list)

BVeq<-filter_firms_and_date(BVeq,2010,end_dates_data,firm_list)


############################################
#ensure values are defined as numerical data
############################################

Scp1[[3]]<-as.numeric(Scp1[[3]])
Scp2[[3]]<-as.numeric(Scp2[[3]])
Scp3[[3]]<-as.numeric(Scp3[[3]])
Scp12[[3]]<-as.numeric(Scp12[[3]])
Asset[[3]]<-as.numeric(Asset[[3]])
CurAsset[[3]]<-as.numeric(CurAsset[[3]])
Rev[[3]]<-as.numeric(Rev[[3]])
Liab[[3]]<-as.numeric(Liab[[3]])
CurLiab[[3]]<-as.numeric(CurLiab[[3]])
RetEarn[[3]]<-as.numeric(RetEarn[[3]])
prc[[3]]<-as.numeric(prc[[3]])
prcCap[[3]]<-as.numeric(prc[[3]])
shrout[[3]]<-as.numeric(shrout[[3]])
NProf[[3]]<-as.numeric(NProf[[3]])
BVeq[[3]]<-as.numeric(BVeq[[3]])
rf[[2]]<-as.numeric(rf[[2]])
MktPrc[[2]]<-as.numeric(MktPrc[[2]])
Inflation[[2]]<-as.numeric(Inflation[[2]])

#Cap Debt at 0
Liab$Liabilities <- pmax(Liab$Liabilities, 0)
CurLiab$CurLiabilities <- pmax(CurLiab$CurLiabilities, 0)


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
  
#############################################
#Get Leverage , Profitability, Liquidity, Equity Buffer
#############################################

#merge Total liabilities, Market Cap, Net profit, Total assets, and Book value equity by date and firm 
Merged_data <- Asset%>%
  full_join(CurAsset, by=c("Firm","Date"))%>%
  full_join(CurLiab, by=c("Firm","Date"))%>%
  full_join(MktCapY, by=c("Firm","Date"))%>%
  full_join(NProf, by=c("Firm","Date"))%>%
  full_join(RetEarn, by=c("Firm","Date"))%>%
  full_join(Liab, by=c("Firm","Date"))%>%
  # sort merged data by firm and then date
  arrange("Firm", "Date")%>%
  # calculate market cap
  mutate(Leverage = Liabilities/MktCap)%>%
  mutate(Profitability = NetProf / Assets)%>%
  mutate(Liquidity = CurAssets / CurLiabilities)%>%
  mutate(EquityBuffer = RetEarning / Assets)

Leverage<-Merged_data%>%
  select(Date, Firm, Leverage)

Profitability<-Merged_data%>%
    select(Date, Firm, Profitability)

Liquidity<-Merged_data%>%
  select(Date, Firm, Liquidity)

EquityBuffer<-Merged_data%>%
  select(Date, Firm, EquityBuffer)

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

Annual_return <- prc %>%
  # Extract year from Date
  mutate(Year = year(Date)) %>%
  # Group by Firm and Year
  group_by(Firm, Year) %>%
  # Calculate the mean of daily log returns
  summarize(avg_daily_log_return = mean(re, na.rm = TRUE),
            Date = max(Date),  # Using the maximum date in the group
            .groups = 'drop') %>%
  # Calculate the approximate annual return from average daily log returns
  mutate(re = (1+avg_daily_log_return)^252 - 1)  %>%
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

Annual_Stoxx_return <- MktPrc %>%
  # Extract year from Date
  mutate(Year = year(Date)) %>%
  # Group by Firm and Year
  group_by(Year) %>%
  # Calculate the mean of daily log returns
  summarize(avg_daily_log_return = mean(rMkt, na.rm = TRUE),
            Date = max(Date),  # Using the maximum date in the group
            .groups = 'drop') %>%
  # Calculate the approximate annual return from average daily log returns
  mutate(rMkt = (1+avg_daily_log_return)^252 - 1)
  
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
  
re<-Annual_return%>%
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

rMkt<-Annual_Stoxx_return%>%
  select(Date, rMkt)

#vMkt<-MktPrcY%>%
  #select(Date, vMkt.annual)

#names(vMkt)[names(vMkt) == "vMkt.annual"] <- "vMkt"

####################################
#Scale variables 
####################################

Rev$Revenue<-Rev$Revenue/10^9
Asset$Assets<-Asset$Assets/10^9
Scp1$Scope1<-Scp1$Scope1/10^6
Scp2$Scope2<-Scp2$Scope2/10^6
Scp3$Scope3<-Scp3$Scope3/10^6
Scp12$Scope12<-Scp12$Scope12/10^6


#########################################
#Scope intensity and scope 1+2 (reported)
#########################################

#Calculate intensity and combined 1 and 2 level
ScpItensity <- Scp1%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp3, by=c("Firm","Date"))%>%
  full_join(Rev, by=c("Firm","Date"))%>%
  mutate(Scope12R=(Scope1 + Scope2))%>%
  mutate(Scp1Int=Scope1 / Revenue)%>%
  mutate(Scp2Int=Scope2 / Revenue)%>%
  mutate(Scp3Int=Scope3 / Revenue)%>%
  mutate(Scp12RInt=Scope12R / Revenue)

ScpItensity12 <- Scp12%>%
  full_join(Rev, by=c("Firm","Date"))%>%
  mutate(Scp12Int=Scope12 / Revenue)

  
ScpItensity<-filter(ScpItensity, Year.x >= 2010)

#Extract intensity and scope 1 + 2 level
Scope12R<-ScpItensity%>%
  select(Date, Firm, Scope12R)

Scp1Int<-ScpItensity%>%
  select(Date, Firm, Scp1Int)

Scp2Int<-ScpItensity%>%
  select(Date, Firm, Scp2Int)

Scp3Int<-ScpItensity%>%
  select(Date, Firm, Scp3Int)

Scp12Int<-ScpItensity12%>%
  select(Date, Firm, Scp12Int)

Scp12RInt<-ScpItensity%>%
  select(Date, Firm, Scp12RInt)

#########################################
#Scope YoY and scope 1+2 (reported)
#########################################

#Calculate Scope YoY
ScpYoY <- Scp1%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp3, by=c("Firm","Date"))%>%
  full_join(Scp12, by=c("Firm","Date"))%>%
  full_join(Rev, by=c("Firm","Date"))%>%
  mutate(Scope12R=(Scope1 + Scope2))%>%
  mutate(Scp12YoY = ifelse(Firm == dplyr::lag(Firm), (Scope12 / dplyr::lag(Scope12) - 1), NA)) %>%
  mutate(Scp12RYoY = ifelse(Firm == dplyr::lag(Firm), (Scope12R / dplyr::lag(Scope12R) - 1), NA))

ScpYoY<-ScpYoY%>%
  filter(!is.na(Scp12YoY))

#Extract YoY

Scp12YoY<-ScpYoY%>%
  select(Date, Firm, Scp12YoY)

Scp12RYoY<-ScpYoY%>%
  select(Date, Firm, Scp12RYoY)

#########################################
#Disclosure + Audit
#########################################

Disclousre_Audit <- Scp1%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp12, by=c("Firm","Date"))%>%
  full_join(Audit, by=c("Firm","Date"))%>%
  full_join(SBTItarget, by=c("Firm","Date"))%>%
  mutate(Scope12R = Scope1 + Scope2 )%>%
  mutate(Disclosure = ifelse(!is.na(Scope12R), 1, 0))%>%
  mutate(DisclosureYoY = ifelse(!is.na(Scope12R) & dplyr::lag(!is.na(Scope12R)), 1, 0))%>%
  mutate(Audit = ifelse(Audit==1, 1, 0))%>%
  mutate(AuditYoY = ifelse(Audit==1 & dplyr::lag(Audit==1), 1, 0))%>%
  mutate(SBTiTarget = ifelse(SBTIdummy==1, 1, 0))%>%
  mutate(SBTiTargetYoY = ifelse(SBTIdummy==1 & dplyr::lag(SBTIdummy==1), 1, 0))

rm(SBTItarget)

Disclosure<-Disclousre_Audit%>%
  select(Date, Firm, Disclosure)%>%
  mutate(Year = year(ymd(Date)))

DisclosureYoY<-Disclousre_Audit%>%
  select(Date, Firm, DisclosureYoY)%>%
  mutate(Year = year(ymd(Date)))

Audit<-Disclousre_Audit%>%
  mutate(Year = year(ymd(Date)))%>%
  select(Date, Firm, Audit)

AuditYoY<-Disclousre_Audit%>%
  mutate(Year = year(ymd(Date)))%>%
  select(Date, Firm, AuditYoY)

SBTiTarget<-Disclousre_Audit%>%
  mutate(Year = year(ymd(Date)))%>%
  select(Date, Firm, SBTiTarget)

SBTiTargetYoY<-Disclousre_Audit%>%
  mutate(Year = year(ymd(Date)))%>%
  select(Date, Firm, SBTiTargetYoY)

  
#########################################
#Remove 2009 data from scope data
#########################################

#Levels
Scp1<- filter(Scp1, Year >= 2010)
Scp2<- filter(Scp2, Year >= 2010)
Scp3<- filter(Scp3, Year >= 2010)
Scp12<- filter(Scp12, Year >= 2010)
Disclosure<- filter(Disclosure, Year >= 2010)
DisclosureYoY<- filter(DisclosureYoY, Year >= 2010)



########################################################
#Combine data for dataset to regression (-DtD set to 0)
########################################################
DtD.results<-Clean_DtD_neg_zero%>%
  select(Date,Firm,DtD)

#Data including scope 3
Reg_dataInclscp3 <- DtD.results%>%
  full_join(Scp1, by=c("Firm","Date"))%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp3, by=c("Firm","Date"))%>%
  full_join(Scp12, by=c("Firm","Date"))%>%
  full_join(Scope12R, by=c("Firm","Date"))%>%
  full_join(Scp1Int, by=c("Firm","Date"))%>%
  full_join(Scp2Int, by=c("Firm","Date"))%>%
  full_join(Scp3Int, by=c("Firm","Date"))%>%
  full_join(Scp12Int, by=c("Firm","Date"))%>%
  full_join(Scp12RInt, by=c("Firm","Date"))%>%
  full_join(Scp12YoY, by=c("Firm","Date"))%>%
  full_join(Scp12RYoY, by=c("Firm","Date"))%>%  
  full_join(SP_clean, by=c("Firm","Date"))%>%
  full_join(Leverage, by=c("Firm","Date"))%>%
  full_join(Asset, by=c("Firm","Date"))%>%
  full_join(Profitability, by=c("Firm","Date"))%>%
  full_join(Liquidity, by=c("Firm","Date"))%>%
  full_join(EquityBuffer, by=c("Firm","Date"))%>%
  full_join(rMkt, by=c("Date"))%>%
  full_join(vMkt, by=c("Date"))%>%
  full_join(Inflation, by=c("Date"))%>%
  full_join(rf, by=c("Date"))%>%
  full_join(Disclosure, by=c("Firm","Date"))%>%
  full_join(DisclosureYoY, by=c("Firm","Date"))%>%
  full_join(AuditYoY, by=c("Firm","Date"))%>%
  full_join(Audit, by=c("Firm","Date"))%>%
  full_join(SBTiTarget, by=c("Firm","Date"))%>%
  full_join(SBTiTargetYoY, by=c("Firm","Date"))%>%
  full_join(Sector, by=c("Firm"))%>%
  full_join(Country, by=c("Firm"))%>%
  # sort merged data by firm and then date
  rename(HYRating = ratingdummy) %>%
  rename(Size = Assets) %>%
  rename(MarketReturn = rMkt) %>%
  rename(MarketVolatility = vMkt) %>%
  rename(RiskFree = Euribor) %>%
  arrange("Date", "Firm")%>%
  distinct()

#Data excluding scope 3
Reg_dataexclscp3<-Reg_dataInclscp3%>%
  select(-Scope3, -Scp3Int)

##################################
#Data selection and final sample
##################################
# Define row names
row_names <- c("Initial sample", 
               "- Financial firms", 
               "- less than 1 year of observations", 
               "- Firms located in Bermuda or Jersey",
               "- Missing values for Control variables", 
               "- Missing values for Scope 1 & 2",
               "Final sample")

# Create a data frame instead of a matrix
Data_Selection <- data.frame(
  "Reason for exclusion" = row_names,
  "Firms excluded" = NA_integer_,  # Initialize with NA of type integer
  "Sample size" = NA_integer_  # Initialize with NA of type integer
)

# Add column names
colnames(Data_Selection) <- c("Reason for exclusion", "Firms excluded", "Sample size")

# Fill values of sample size
##Intial sample
Data_Selection[1, 3] <- 600

##Finacial firms
Data_Selection[2, 3] <- 450

##Less than 1 year of observations
load("firm_list.Rdata")
Data_Selection[3, 3] <- length(unique(firm_list$Firm))

##Firms located in Bermuda or Jersey
Reg_dataexclscp3 <- Reg_dataexclscp3 %>%
  filter(Country!="Bermuda" & Country!="Jersey")
Data_Selection[4, 3] <- length(unique(Reg_dataexclscp3$Firm))

##Missing values for Control variables
  Reg_dataexclscp3 <- Reg_dataexclscp3 %>%
    filter(!is.na(Leverage) &!is.na(Size) & !is.na(Profitability) & !is.na(Liquidity)& !is.na(EquityBuffer))
    Data_Selection[5, 3] <- length(unique(Reg_dataexclscp3$Firm))
  
##Missing values for Scope 1 & 2 
  Reg_dataexclscp3 <- Reg_dataexclscp3 %>%
  filter(!is.na(DtD) & !is.na(Scope12) & !is.na(Scp12Int))
  Data_Selection[6, 3] <- length(unique(Reg_dataexclscp3$Firm))
  
##Final sample
Data_Selection[7, 3] <- length(unique(Reg_dataexclscp3$Firm))


#Fill out firms excluded
for (i in 2:(nrow(Data_Selection) - 1)) {
  Data_Selection[i, 2] <- Data_Selection[i - 1, 3] - Data_Selection[i, 3]
}

# Print the matrix
print(Data_Selection)

Data_Selection[] <- lapply(Data_Selection, function(x) if(is.numeric(x)) as.integer(x) else x)

latex_code1 <- print(xtable(Data_Selection), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE, digits = 0)

##Variables
#Date
#Firm
#DtD
#Scope1
#Scope2
#Scope3
#Scope12R
#Scp1Int
#Scp2Int
#Scp3Int
#Scp12RInt
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

##################################
#Winzorize firm control variables
##################################

Data_Wins<-Reg_dataexclscp3 %>%
  mutate(Leverage = Winsorize(Leverage, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Size = Winsorize(Size, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Profitability = Winsorize(Profitability, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Liquidity = Winsorize(Liquidity, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(EquityBuffer = Winsorize(EquityBuffer, probs = c(0.01, 0.99), na.rm = TRUE))
  
##################################
#Extract regression data
##################################

Reg_Data<-Reg_dataexclscp3

save(Reg_Data,file="Reg_Data.Rdata")
save(Data_Wins,file="Reg_Data_Wins.Rdata")


########################################################
#Combine data for dataset to regression (Remove -DtD)
########################################################

DtD.results2<-Clean_DtD_neg_rm%>%
  select(Date,Firm,DtD)

#Data including scope 3
Reg_dataInclscp31 <- DtD.results2%>%
  full_join(Scp1, by=c("Firm","Date"))%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp3, by=c("Firm","Date"))%>%
  full_join(Scp12, by=c("Firm","Date"))%>%
  full_join(Scope12R, by=c("Firm","Date"))%>%
  full_join(Scp1Int, by=c("Firm","Date"))%>%
  full_join(Scp2Int, by=c("Firm","Date"))%>%
  full_join(Scp3Int, by=c("Firm","Date"))%>%
  full_join(Scp12Int, by=c("Firm","Date"))%>%
  full_join(Scp12RInt, by=c("Firm","Date"))%>%
  full_join(Scp12YoY, by=c("Firm","Date"))%>%
  full_join(Scp12RYoY, by=c("Firm","Date"))%>%  
  full_join(Audit, by=c("Firm","Date"))%>%
  full_join(SBTItarget, by=c("Firm","Date"))%>%
  full_join(SP_clean, by=c("Firm","Date"))%>%
  full_join(Leverage, by=c("Firm","Date"))%>%
  full_join(Asset, by=c("Firm","Date"))%>%
  full_join(Profitability, by=c("Firm","Date"))%>%
  full_join(Liquidity, by=c("Firm","Date"))%>%
  full_join(EquityBuffer, by=c("Firm","Date"))%>%
  full_join(rMkt, by=c("Date"))%>%
  full_join(vMkt, by=c("Date"))%>%
  full_join(Inflation, by=c("Date"))%>%
  full_join(rf, by=c("Date"))%>%
  full_join(Sector, by=c("Firm"))%>%
  full_join(Country, by=c("Firm"))%>%
  # sort merged data by firm and then date
  rename(SBTiTarget = SBTIdummy) %>%
  rename(HYRating = ratingdummy) %>%
  rename(Size = Assets) %>%
  rename(MarketReturn = rMkt) %>%
  rename(MarketVolatility = vMkt) %>%
  rename(RiskFree = Euribor) %>%
  arrange("Date", "Firm")%>%
  distinct()


#Data excluding scope 3
Reg_dataexclscp31<-Reg_dataInclscp31%>%
  select(-Scope3, -Scp3Int)

#Exclusions

##Firms located in Bermuda
Reg_dataexclscp31 <- Reg_dataexclscp31 %>%
  filter(Country!="Bermuda" & Country!="Jersey")

##Missing values for DtD
Reg_dataexclscp31 <- Reg_dataexclscp31 %>%
  filter(!is.na(DtD))

##Missing values for Control variables
Reg_dataexclscp31 <- Reg_dataexclscp31 %>%
  filter(!is.na(Leverage) &!is.na(Size) & !is.na(Profitability) & !is.na(Liquidity)& !is.na(EquityBuffer))

##Missing values for Scope 1 & 2 
Reg_dataexclscp31 <- Reg_dataexclscp31 %>%
  filter(!is.na(Scope12) & !is.na(Scp12Int))

##Variables
#Date
#Firm
#DtD
#Scope1
#Scope2
#Scope3
#Scope12R
#Scp1Int
#Scp2Int
#Scp3Int
#Scp12RInt
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

##################################
#Winzorize firm control variables
##################################

Data_Wins1<-Reg_dataexclscp31 %>%
  mutate(Leverage = Winsorize(Leverage, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Size = Winsorize(Size, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Profitability = Winsorize(Profitability, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Liquidity = Winsorize(Liquidity, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(EquityBuffer = Winsorize(EquityBuffer, probs = c(0.01, 0.99), na.rm = TRUE))

##################################
#Extract regression data
##################################

Reg_Data1<-Reg_dataexclscp31

save(Reg_Data1,file="Reg_Data1.Rdata")
save(Data_Wins1,file="Reg_Data_Wins1.Rdata")

########################################################
#Combine data for dataset to regression (DtD wip)
########################################################

DtD.results3<-DtD_wip%>%
  select(Date,Firm,DtD)

#Data including scope 3
Reg_dataInclscp32 <- DtD.results3%>%
  full_join(Scp1, by=c("Firm","Date"))%>%
  full_join(Scp2, by=c("Firm","Date"))%>%
  full_join(Scp3, by=c("Firm","Date"))%>%
  full_join(Scp12, by=c("Firm","Date"))%>%
  full_join(Scope12R, by=c("Firm","Date"))%>%
  full_join(Scp1Int, by=c("Firm","Date"))%>%
  full_join(Scp2Int, by=c("Firm","Date"))%>%
  full_join(Scp3Int, by=c("Firm","Date"))%>%
  full_join(Scp12Int, by=c("Firm","Date"))%>%
  full_join(Scp12RInt, by=c("Firm","Date"))%>%
  full_join(Scp12YoY, by=c("Firm","Date"))%>%
  full_join(Scp12RYoY, by=c("Firm","Date"))%>%  
  full_join(Audit, by=c("Firm","Date"))%>%
  full_join(SBTItarget, by=c("Firm","Date"))%>%
  full_join(SP_clean, by=c("Firm","Date"))%>%
  full_join(Leverage, by=c("Firm","Date"))%>%
  full_join(Asset, by=c("Firm","Date"))%>%
  full_join(Profitability, by=c("Firm","Date"))%>%
  full_join(Liquidity, by=c("Firm","Date"))%>%
  full_join(EquityBuffer, by=c("Firm","Date"))%>%
  full_join(rMkt, by=c("Date"))%>%
  full_join(vMkt, by=c("Date"))%>%
  full_join(Inflation, by=c("Date"))%>%
  full_join(rf, by=c("Date"))%>%
  full_join(Sector, by=c("Firm"))%>%
  full_join(Country, by=c("Firm"))%>%
  # sort merged data by firm and then date
  rename(SBTiTarget = SBTIdummy) %>%
  rename(HYRating = ratingdummy) %>%
  rename(Size = Assets) %>%
  rename(MarketReturn = rMkt) %>%
  rename(MarketVolatility = vMkt) %>%
  rename(RiskFree = Euribor) %>%
  arrange("Date", "Firm")%>%
  distinct()

#Data excluding scope 3
Reg_dataexclscp32<-Reg_dataInclscp32%>%
  select(-Scope3, -Scp3Int)

#Exclusions

##Firms located in Bermuda
Reg_dataexclscp32 <- Reg_dataexclscp32 %>%
  filter(Country!="Bermuda" & Country!="Jersey")

##Missing values for DtD
Reg_dataexclscp32 <- Reg_dataexclscp32 %>%
  filter(!is.na(DtD))

##Missing values for Control variables
Reg_dataexclscp32 <- Reg_dataexclscp32 %>%
  filter(!is.na(Leverage) &!is.na(Size) & !is.na(Profitability) & !is.na(Liquidity)& !is.na(EquityBuffer))

##Missing values for Scope 1 & 2 
Reg_dataexclscp32 <- Reg_dataexclscp32 %>%
  filter(!is.na(Scope12) & !is.na(Scp12Int))

##Variables
#Date
#Firm
#DtD
#Scope1
#Scope2
#Scope3
#Scope12R
#Scp1Int
#Scp2Int
#Scp3Int
#Scp12RInt
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

##################################
#Winzorize firm control variables
##################################

Data_Wins2<-Reg_dataexclscp32 %>%
  mutate(Leverage = Winsorize(Leverage, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Size = Winsorize(Size, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Profitability = Winsorize(Profitability, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(Liquidity = Winsorize(Liquidity, probs = c(0.01, 0.99), na.rm = TRUE)) %>%
  mutate(EquityBuffer = Winsorize(EquityBuffer, probs = c(0.01, 0.99), na.rm = TRUE))

##################################
#Extract regression data
##################################

Reg_Data2<-Reg_dataexclscp32

save(Reg_Data2,file="Reg_Data2.Rdata")
save(Data_Wins2,file="Reg_Data_Wins2.Rdata")
