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

##############
#Data frame
##############

#Table of data not winsorized
Summary_table1_data<-Data_Wins%>%
  select(Scp12Int, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , RiskFree)%>%
  mutate(lSize=log(Size))%>%
  select(-Size)


#Rename name variables for tables


# Creating summary statistics table
Summary_table1 <- Summary_table1_data %>%
  summarise(across(everything(), list(
    Mean = ~mean(., na.rm = TRUE)
  ))) %>%
  pivot_longer(
    cols = everything(),
    names_to = c("Variable", ".value"),
    names_sep = "_"
  )

################
#Models
################

Data_Wins <- Data_Wins %>%
  mutate(lSize = log(Size)) %>%
  select(-Size)


#Prepare data for models
Data_Wins$Audit<-as.factor(Data_Wins$Audit)
Data_Wins$SBTiTarget<-as.factor(Data_Wins$SBTiTarget)
Data_Wins$HYRating<-as.factor(Data_Wins$HYRating)

Data_Wins$Year.y.y<-as.factor(Data_Wins$Year.y.y)
Data_Wins$Sector_number<-as.factor(Data_Wins$Sector_number)
Data_Wins$Country_number<-as.factor(Data_Wins$Country_number)

Data_Wins$sector_country <- interaction(Data_Wins$Sector_number, Data_Wins$Country_number)
Data_Wins$sector_country <- as.factor(Data_Wins$sector_country)

DataHY<-Data_Wins%>%
  filter(!is.na(HYRating)) 

#Base model
Base <- lm(DtD ~ Scp12Int + Leverage + lSize +Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility  + RiskFree  + Year.y.y + sector_country, data = Data_Wins)

#Audit model
Audit <- lm(DtD ~ Scp12Int*Audit + Leverage + lSize + Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = Data_Wins)

#SBTi model
SBTi <- lm(DtD ~ Scp12Int*SBTiTarget + Leverage + lSize + Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = Data_Wins)

#IG vs HY model
HY <- lm(DtD ~ Scp12Int*HYRating + Leverage + lSize + Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = DataHY)


############################
#Combine model coefficients
############################

# Step 1: Extract coefficients and calculate averages for group and time fixed effects
extract_and_average_fixed_effects <- function(model, group_prefix, time_prefix) {
  # Get all coefficients
  coeffs <- coef(model)
  
  # Identify coefficients related to group (sector-country) and time (year)
  group_coeffs <- coeffs[grepl(group_prefix, names(coeffs))]
  time_coeffs <- coeffs[grepl(time_prefix, names(coeffs))]
  
  # Calculate average for group and time fixed effects
  avg_group <- mean(group_coeffs, na.rm = TRUE)
  avg_time <- mean(time_coeffs, na.rm = TRUE)
  
  return(list(coeffs = coeffs, avg_group = avg_group, avg_time = avg_time))
}

# Extract coefficients and calculate averages for all models
base_effects <- extract_and_average_fixed_effects(Base, "sector_country", "Year.y.y")
audit_effects <- extract_and_average_fixed_effects(Audit, "sector_country", "Year.y.y")
sbti_effects <- extract_and_average_fixed_effects(SBTi, "sector_country", "Year.y.y")
hy_effects <- extract_and_average_fixed_effects(HY, "sector_country", "Year.y.y")


# Compute the mean values
mean_vals <- Summary_table1_data %>%
  summarise(across(everything(), mean, na.rm = TRUE)) %>%
  as.list()

# Function to calculate the product of coefficients and means
calc_effects <- function(coeffs, mean_vals) {
  effect_vals <- numeric()
  
  # Calculate for intercept
  effect_vals["(Intercept)"] <- coeffs["(Intercept)"]
  
  # Multiply non-dummy variables by their mean values
  for (var in names(mean_vals)) {
    if (var %in% names(coeffs)) {
      effect_vals[var] <- coeffs[var] * mean_vals[[var]]
    }
  }
  
  # Add dummy variable coefficients directly without modification
  dummy_vars <- names(coeffs)[grepl("Audit|SBTiTarget|HYRating", names(coeffs))]
  for (dummy in dummy_vars) {
    effect_vals[dummy] <- coeffs[dummy]
  }
  
  return(effect_vals)
}
# Function to calculate the product of coefficients and means with consideration for interactions
calc_effects_with_interaction <- function(coeffs, mean_vals, interaction_term) {
  effect_vals <- numeric()
  
  # Calculate for intercept
  effect_vals["(Intercept)"] <- coeffs["(Intercept)"]
  
  # Multiply non-dummy variables by their mean values
  for (var in names(mean_vals)) {
    if (var %in% names(coeffs)) {
      effect_vals[var] <- coeffs[var] * mean_vals[[var]]
    }
  }
  
  # Handle interaction terms involving Scp12Int
  if (interaction_term %in% names(coeffs)) {
    combined_effect <- coeffs["Scp12Int"] + coeffs[interaction_term]
    effect_vals["Scp12Int"] <- combined_effect * mean_vals[["Scp12Int"]]
  }
  
  # Add dummy variable coefficients directly without modification
  dummy_vars <- names(coeffs)[grepl("Audit|SBTiTarget|HYRating", names(coeffs))]
  for (dummy in dummy_vars) {
    effect_vals[dummy] <- coeffs[dummy]
  }
  
  return(effect_vals)
}

# Step 3: Calculate the effect values for each model
base_calc <- calc_effects(base_effects$coeffs, mean_vals)
audit_calc <- calc_effects_with_interaction(audit_effects$coeffs, mean_vals, "Scp12Int:Audit1")
sbti_calc <- calc_effects_with_interaction(sbti_effects$coeffs, mean_vals, "Scp12Int:SBTiTarget1")
hy_calc <- calc_effects_with_interaction(hy_effects$coeffs, mean_vals, "Scp12Int:HYRating1")

# Step 4: Add group and time effects
base_calc["Group Effects"] <- base_effects$avg_group
audit_calc["Group Effects"] = audit_effects$avg_group
sbti_calc["Group Effects"] = sbti_effects$avg_group
hy_calc["Group Effects"] = hy_effects$avg_group

base_calc["Time Fixed Effects"] <- base_effects$avg_time
audit_calc["Time Fixed Effects"] = audit_effects$avg_time
sbti_calc["Time Fixed Effects"] = sbti_effects$avg_time
hy_calc["Time Fixed Effects"] = hy_effects$avg_time

# Create a vector of mean values with consistent length
mean_vals_vector <- c(
  mean_vals,  # The means of numerical variables
  "Group Effects" = NA,  # Add NA placeholders for non-numeric variables
  "Time Fixed Effects" = NA
)

# If the coefficients include any dummy variables, append them as NA placeholders
dummy_vars <- c("Audit", "SBTiTarget", "HYRating")
mean_vals_vector[dummy_vars] <- NA

# Ensure that the vector has the same length and ordering as base_calc
mean_vals_vector <- mean_vals_vector[names(base_calc)]

base_df <- data.frame(Variable = names(base_calc), Effect = unname(base_calc), stringsAsFactors = FALSE)
audit_df <- data.frame(Variable = names(audit_calc), Effect = unname(audit_calc), stringsAsFactors = FALSE)
sbti_df <- data.frame(Variable = names(sbti_calc), Effect = unname(sbti_calc), stringsAsFactors = FALSE)
hy_df <- data.frame(Variable = names(hy_calc), Effect = unname(hy_calc), stringsAsFactors = FALSE)
