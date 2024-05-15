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
  select(DtD, Scope12, Scp12Int, Scp12YoY, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree, Year.y.y, Sector_number, Country_number, Country, Firm)

Data <- Data %>%
  mutate(year_group = case_when(
    Year.y.y >= 2010 & Year.y.y <= 2015 ~ 1,
    Year.y.y >= 2016 & Year.y.y <= 2022 ~ 2,
  ))

max_year_group <- max(Data$year_group, na.rm = TRUE)

#Prepare data for the first model
Data$Sector_number<-as.factor(Data$Sector_number)
Data$Country_number<-as.factor(Data$Country_number)
Data$Year.y.y<-as.factor(Data$Year.y.y)

Data$sector_country <- interaction(Data$Sector_number, Data$Country_number)
Data$sector_country <- as.factor(Data$sector_country)


####################
#Models for Level
####################

# Initialize an output list
results_df <- data.frame(Sector = character(), Variable = character(), Coefficient = numeric(), Significance = character())
outputs <- list()

# Loop through each Country
for (i in 1:max_year_group) {
  
  YearGroupData <- filter(Data, year_group == i)
  
  # Model 1
  model1 <- lm(DtD ~ Scope12 + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + Inflation + RiskFree + Year.y.y + sector_country, data = YearGroupData)
  
  # Model 2
  model2 <- lm(DtD ~ Scp12Int + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + Inflation + RiskFree + Year.y.y + sector_country, data = YearGroupData)
  
  # Model 3
  model3 <- lm(DtD ~ Scp12YoY + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + Inflation + RiskFree + Year.y.y + sector_country, data = YearGroupData)
  
  # Adjust standard errors for clustering at the firm level
  model1_results <- coeftest(model1, vcov = vcovCL(model1, cluster = ~Firm, data = Data))
  model2_results <- coeftest(model2, vcov = vcovCL(model2, cluster = ~Firm, data = Data))
  model3_results <- coeftest(model3, vcov = vcovCL(model3, cluster = ~Firm, data = Data))
  
  # Capture output of stargazer
  output <- capture.output(
    stargazer(model1, model2, model3, type = "latex",
              coef=list(model1_results[,1], model2_results[,1], model3_results[,1]),
              se=list(model1_results[,2], model2_results[,2], model3_results[,2]),
              t=list(model1_results[,3], model2_results[,3], model3_results[,3]),
              p=list(model1_results[,4], model2_results[,4], model3_results[,4]), 
              title = paste("Regression Models Comparison for Year Group", i),
              omit = c("sector_country"), 
              add.lines = list(c("Country fixed-effects", "Yes", "Yes", "Yes"), 
                               c("Time fixed-effects", "Yes", "Yes", "Yes")), 
              omit.stat = c("adj.rsq", "ser", "f"), 
              align = TRUE)
  )
  
  # Save output in a text file
  writeLines(output, paste0("9.1_","Year_Group_", i, "_models.tex"))
  
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
        Year = i,
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
    Scope12 = ifelse(!is.na(Coefficient_Scope12), paste0(sprintf("%.2f", Coefficient_Scope12), Significance_Scope12), NA),
    Scp12Int = ifelse(!is.na(Coefficient_Scp12Int), paste0(sprintf("%.2f", Coefficient_Scp12Int), Significance_Scp12Int), NA),
    Scp12YoY = ifelse(!is.na(Coefficient_Scp12YoY), paste0(sprintf("%.2f", Coefficient_Scp12YoY), Significance_Scp12YoY), NA)
  ) %>%
  select(Year, Scope12, Scp12Int, Scp12YoY)

# Print to check final format
print(results_wide)

summary<- print(xtable(results_wide), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

# Save output in a text file
writeLines(summary, paste0("9.1_Summary.tex"))

