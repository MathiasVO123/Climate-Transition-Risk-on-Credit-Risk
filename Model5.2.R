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
library(ggplot2)
library(extrafont)
library(xtable)
library(plm)
library(purrr)
library(lfe)
library(pracma)
library(lmtest)
library(car)
library(gridExtra)
library(stargazer)

#step 2+2 only necessary the first time
#font_path <- "C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code"
#font_import(paths = font_path, prompt = FALSE)

loadfonts()

#removing all previous variables 
rm(list=ls())

#setting WD to code folder  
setwd("C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#load functions 
source("Thesis functions.R")

#Load Data
load("Reg_Data.Rdata")
load("Reg_Data_Wins.Rdata")

Data<-Data_Wins%>%
  select(DtD, Scope12, Scp12Int, Scp12YoY, HYRating, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree, Year.y.y, Sector_number, Country_number, Firm)

#Prepare data for the first model
Data$Year.y.y<-as.factor(Data$Year.y.y)
Data$Sector_number<-as.factor(Data$Sector_number)
Data$Country_number<-as.factor(Data$Country_number)

Data$sector_country <- interaction(Data$Sector_number, Data$Country_number)
Data$sector_country <- as.factor(Data$sector_country)

####################
#Models for HY
####################

#Data sets 
DataHY<-Data%>%
  filter(HYRating==1)

DataIG<-Data%>%
  filter(HYRating==0)

summary(DataHY)
summary(DataIG)


####HY models
#Model 1
model1 <- lm(DtD ~ Scope12 + Leverage + log(Size)+ Profitability+ Liquidity + EquityBuffer+ MarketReturn + MarketVolatility + RiskFree + Year.y.y + sector_country, data = DataHY)
summary(model1)

#Model 2
model2 <- lm(DtD ~ Scp12Int + Leverage + log(Size)+ Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = DataHY)
summary(model2)

#Model 3
model3 <- lm(DtD ~ Scp12YoY + Leverage + log(Size)+ Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = DataHY)
summary(model3)

# Adjust standard errors for clustering at the firm level
model1_results <- coeftest(model1, vcov = vcovCL(model1, cluster = ~Firm, data = DataHY))
model2_results <- coeftest(model2, vcov = vcovCL(model2, cluster = ~Firm, data = DataHY))
model3_results <- coeftest(model3, vcov = vcovCL(model3, cluster = ~Firm, data = DataHY))

# Use stargazer to create the table, specifying each parameter manually
stargazer(model1, model2, model3, type = "latex",
          coef=list(model1_results[,1], model2_results[,1], model3_results[,1]),
          se=list(model1_results[,2], model2_results[,2], model3_results[,2]),
          t=list(model1_results[,3], model2_results[,3], model3_results[,3]),
          p=list(model1_results[,4], model2_results[,4], model3_results[,4]),  
          title = "Regression Models Comparison",
          omit = c("Year.y.y", "sector_country"), 
          add.lines = list(c("Country&Sector fixed-effects", "Yes", "Yes", "Yes"), 
                           c("Time fixed-effects", "Yes", "Yes", "Yes")),
          omit.stat = c("adj.rsq", "ser", "f"), 
          align = TRUE)

####IG models
#Model 1
model12 <- lm(DtD ~ Scope12 + Leverage + log(Size)+ Profitability+ Liquidity + EquityBuffer+ MarketReturn + MarketVolatility + RiskFree + Year.y.y + sector_country, data = DataIG)
summary(model1)

#Model 2
model22 <- lm(DtD ~ Scp12Int + Leverage + log(Size)+ Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = DataIG)
summary(model2)

#Model 3
model32 <- lm(DtD ~ Scp12YoY + Leverage + log(Size)+ Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = DataIG)
summary(model3)

# Adjust standard errors for clustering at the firm level
model12_results <- coeftest(model12, vcov = vcovCL(model1, cluster = ~Firm, data = DataIG))
model22_results <- coeftest(model22, vcov = vcovCL(model2, cluster = ~Firm, data = DataIG))
model32_results <- coeftest(model32, vcov = vcovCL(model3, cluster = ~Firm, data = DataIG))

# Use stargazer to create the table, specifying each parameter manually
stargazer(model12, model22, model32, type = "latex",
          coef=list(model12_results[,1], model22_results[,1], model32_results[,1]),
          se=list(model12_results[,2], model22_results[,2], model32_results[,2]),
          t=list(model12_results[,3], model22_results[,3], model32_results[,3]),
          p=list(model12_results[,4], model22_results[,4], model32_results[,4]),  
          title = "Regression Models Comparison",
          omit = c("Year.y.y", "sector_country"), 
          add.lines = list(c("Country&Sector fixed effects", "Yes", "Yes", "Yes"), 
                           c("Time fixed-effects", "Yes", "Yes", "Yes")),
          omit.stat = c("adj.rsq", "ser", "f"), 
          align = TRUE)
