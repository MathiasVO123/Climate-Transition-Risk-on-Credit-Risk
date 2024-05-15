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
library(broom)

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
  select(DtD, Scope12, Scp12Int, Scp12YoY, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree, Year.y.y, Sector_number, Country_number, Country, Firm)


# Create a mapping from Country_number to Country
country_mapping <- Data %>%
  select(Country_number, Country) %>%
  distinct()

#Remove Countries with less than 100 observations
Countries_with_few_observations <- Data %>%
  group_by(Country_number) %>%
  filter(n() < 100) %>%
  summarize(count = n(), .groups = 'drop') %>%
  pull(Country_number)

# Data frame where Countries have at least 100 observations
Data_filtered <- Data %>%
  filter(Country_number %in% setdiff(unique(Data$Country_number), Countries_with_few_observations))

# Output the Countries removed
Countries_removed <- setdiff(unique(Data$Country_number), unique(Data_filtered$Country_number))
Countries_removed

#Prepare data for the first model
Data$Year.y.y<-as.factor(Data$Year.y.y)
Data$Sector_number<-as.factor(Data$Sector_number)

####################
#Models for Level
####################

# Determine the maximum Country number dynamically
max_Country <- max(Data$Country_number, na.rm = TRUE)

# Initialize an output list
results_df <- data.frame(Sector = character(), Variable = character(), Coefficient = numeric(), Significance = character())
outputs <- list()

# Loop through each Country
for (i in 1:max_Country) {
  
  if(i %in% Countries_removed) {
    next  # Skip the iteration if country number is in Countries_removed
  }
  
  CountryData <- filter(Data, Country_number == i)
  country_name <- country_mapping$Country[country_mapping$Country_number == i]
  
  
  # Model 1
  model1 <- lm(DtD ~ Scope12 + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + RiskFree + Year.y.y + Sector_number, data = CountryData)
  
  # Model 2
  model2 <- lm(DtD ~ Scp12Int + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + RiskFree + Year.y.y + Sector_number, data = CountryData)
  
  # Model 3
  model3 <- lm(DtD ~ Scp12YoY + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + RiskFree + Year.y.y + Sector_number, data = CountryData)
  
  # Adjust standard errors for clustering at the firm level
  model1_results <- coeftest(model1, vcov = vcovCL(model1, cluster = ~Firm, data = CountryData))
  model2_results <- coeftest(model2, vcov = vcovCL(model2, cluster = ~Firm, data = CountryData))
  model3_results <- coeftest(model3, vcov = vcovCL(model3, cluster = ~Firm, data = CountryData))
  
  # Capture output of stargazer
  output <- capture.output(
    stargazer(model1, model2, model3, type = "latex",
              coef=list(model1_results[,1], model2_results[,1], model3_results[,1]),
              se=list(model1_results[,2], model2_results[,2], model3_results[,2]),
              t=list(model1_results[,3], model2_results[,3], model3_results[,3]),
              p=list(model1_results[,4], model2_results[,4], model3_results[,4]), 
              title = paste("Regression Models Comparison for",country_name),
              omit = c("Year.y.y", "Sector_number"), 
              add.lines = list(c("Sector fixed-effects", "Yes", "Yes", "Yes"), 
                               c("Time fixed-effects", "Yes", "Yes", "Yes")),
              omit.stat = c("adj.rsq", "ser", "f"), 
              align = TRUE)
  )
  
  # Save output in a text file
  writeLines(output, paste0(i,country_name, "_models.tex"))
  
  # Optionally store in list to check in R environment
  outputs[[i]] <- output
  
  # Extract coefficients and p-values for each model separately
  tidy1 <- tidy(model1_results)
  tidy2 <- tidy(model2_results)
  tidy3 <- tidy(model3_results)
  
  # Combine and filter the tidy data
  coef_df <- bind_rows(tidy1, tidy2, tidy3)
  coef_df <- coef_df %>%
    filter(term %in% c("Scope12", "Scp12Int", "Scp12YoY")) %>%
    mutate(Significance = case_when(
      p.value < 0.01 ~ "***",
      p.value < 0.05  ~ "**",
      p.value < 0.10  ~ "*",
      TRUE            ~ ""
    ))
  
  # Append results to the summary dataframe
  for (var in c("Scope12", "Scp12Int", "Scp12YoY")) {
    if (any(coef_df$term == var)) {  # Check if variable exists in the model output
      results_df <- rbind(results_df, data.frame(
        Country = country_name,
        Variable = var,
        Coefficient = coef_df$estimate[coef_df$term == var],
        Significance = coef_df$Significance[coef_df$term == var]
      ))
    }
  }
}

# Pivot the data to wide format correctly
results_wide <- results_df %>%
  pivot_wider(
    names_from = Variable, 
    values_from = c(Coefficient, Significance),
    names_sep = "_"  # This will create names like Scope12_Coefficient
  )

# Combine coefficient and significance into one string
results_wide <- results_wide %>%
  mutate(
    Scope12 = ifelse(!is.na(Coefficient_Scope12), paste0(sprintf("%.3f", Coefficient_Scope12), Significance_Scope12), NA),
    Scp12Int = ifelse(!is.na(Coefficient_Scp12Int), paste0(sprintf("%.3f", Coefficient_Scp12Int), Significance_Scp12Int), NA),
    Scp12YoY = ifelse(!is.na(Coefficient_Scp12YoY), paste0(sprintf("%.3f", Coefficient_Scp12YoY), Significance_Scp12YoY), NA)
  ) %>%
  select(Country, Scope12, Scp12Int, Scp12YoY)

# Print to check final format
print(results_wide)

summary<- print(xtable(results_wide), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

# Save output in a text file
writeLines(summary, paste0("7.1_Summary.tex"))
