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
  select(DtD, Scope12, Scp12Int, Scp12YoY, Disclosure, DisclosureYoY, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree, Year.y.y, Sector_number, Country_number, Country, Firm)


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
Data$Disclosure<-as.factor(Data$Disclosure)
Data$DisclosureYoY<-as.factor(Data$DisclosureYoY)

####################
#Models for Level
####################

# Determine the maximum Country number dynamically
max_Country <- max(Data$Country_number, na.rm = TRUE)

# Initialize an output list
results_df1 <- data.frame(Sector = character(), Variable = character(), Coefficient = numeric(), Significance = character())
results_df2 <- data.frame(Sector = character(), Variable = character(), Coefficient = numeric(), Significance = character())
results_df3 <- data.frame(Sector = character(), Variable = character(), Coefficient = numeric(), Significance = character())
outputs <- list()

# Loop through each Country
for (i in 1:max_Country) {
  
  if(i %in% Countries_removed) {
    next  # Skip the iteration if country number is in Countries_removed
  }
  
  CountryData <- filter(Data, Country_number == i)
  country_name <- country_mapping$Country[country_mapping$Country_number == i]
  
  
  # Model 1
  model1 <- lm(DtD ~ Scope12*Disclosure + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + RiskFree + Year.y.y + Sector_number, data = CountryData)
  
  # Model 2
  model2 <- lm(DtD ~ Scp12Int*Disclosure + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + RiskFree + Year.y.y + Sector_number, data = CountryData)
  
  # Model 3
  model3 <- lm(DtD ~ Scp12YoY*DisclosureYoY + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + RiskFree + Year.y.y + Sector_number, data = CountryData)
  
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
  writeLines(output, paste0("7.2_",country_name, "_models.tex"))
  
  # Optionally store in list to check in R environment
  outputs[[i]] <- output

  # Extract coefficients and p-values for each model separately
  tidy1 <- tidy(model1_results)
  tidy2 <- tidy(model2_results)
  tidy3 <- tidy(model3_results)
  
  # Combine and filter the tidy data
  coef_df1 <- tidy1 %>%
    filter(term %in% c("Scope12", "Disclosure1", "Scope12:Disclosure1")) %>%
    mutate(Significance = case_when(
      p.value < 0.01 ~ "***",
      p.value < 0.05 ~ "**",
      p.value < 0.10 ~ "*",
      TRUE           ~ ""
    ))
  
  # Combine and filter the tidy data
  coef_df2 <- bind_rows(tidy2)
  coef_df2 <- coef_df2 %>%
    filter(term %in% c("Scp12Int", "Disclosure1", "Scp12Int:Disclosure1")) %>%
    mutate(Significance = case_when(
      p.value < 0.01 ~ "***",
      p.value < 0.05  ~ "**",
      p.value < 0.10  ~ "*",
      TRUE            ~ ""
    ))
  
  # Combine and filter the tidy data
  coef_df3 <- bind_rows(tidy3)
  coef_df3 <- coef_df3 %>%
    filter(term %in% c("Scp12YoY", "DisclosureYoY1","Scp12YoY:DisclosureYoY1")) %>%
    mutate(Significance = case_when(
      p.value < 0.01 ~ "***",
      p.value < 0.05  ~ "**",
      p.value < 0.10  ~ "*",
      TRUE            ~ ""
    ))
  
  # Append results to the summary dataframe
  for (var in c("Scope12", "Disclosure1","Scope12:Disclosure1")) {
    if (any(coef_df1$term == var)) {  # Check if variable exists in the model output
      results_df1 <- rbind(results_df1, data.frame(
        Country = country_name,
        Variable = var,
        Coefficient = coef_df1$estimate[coef_df1$term == var],
        Significance = coef_df1$Significance[coef_df1$term == var]
      ))
    }
  }
  
  # Append results to the summary dataframe
  for (var in c("Scp12Int", "Disclosure1", "Scp12Int:Disclosure1")) {
    if (any(coef_df2$term == var)) {  # Check if variable exists in the model output
      results_df2 <- rbind(results_df2, data.frame(
        Country = country_name,
        Variable = var,
        Coefficient = coef_df2$estimate[coef_df2$term == var],
        Significance = coef_df2$Significance[coef_df2$term == var]
      ))
    }
  }
  
# Append results to the summary dataframe
for (var in c("Scp12YoY", "DisclosureYoY1","Scp12YoY:DisclosureYoY1")) {
  if (any(coef_df3$term == var)) {  # Check if variable exists in the model output
    results_df3 <- rbind(results_df3, data.frame(
      Country = country_name,
      Variable = var,
      Coefficient = coef_df3$estimate[coef_df3$term == var],
      Significance = coef_df3$Significance[coef_df3$term == var]
    ))
  }
}

}

######################################
#Get summary tables
######################################

#### Table 1
# Pivot the data to wide format correctly
results_df1 <- results_df1 %>%
  pivot_wider(
    names_from = Variable, 
    values_from = c(Coefficient, Significance),
    names_sep = "_"  # This will create names like Scope12_Coefficient
  )


#### Table 1
# Pivot the data to wide format correctly
results_df2 <- results_df2 %>%
  pivot_wider(
    names_from = Variable, 
    values_from = c(Coefficient, Significance),
    names_sep = "_"  # This will create names like Scope12_Coefficient
  )

#### Table 1
# Pivot the data to wide format correctly
results_df3 <- results_df3 %>%
  pivot_wider(
    names_from = Variable, 
    values_from = c(Coefficient, Significance),
    names_sep = "_"  # This will create names like Scope12_Coefficient
  )


###Table 1
# Combine coefficient and significance into one string
results_df1 <- results_df1 %>%
  mutate(
    Scope12 = ifelse(!is.na(Coefficient_Scope12), paste0(sprintf("%.3f", Coefficient_Scope12), Significance_Scope12), NA),
    `Scope12:Disclosure1` = ifelse(!is.na(`Coefficient_Scope12:Disclosure1`), paste0(sprintf("%.3f", `Coefficient_Scope12:Disclosure1`), `Significance_Scope12:Disclosure1`), NA),
    Disclosure1 = ifelse(!is.na(Coefficient_Disclosure1), paste0(sprintf("%.3f", Coefficient_Disclosure1), Significance_Disclosure1), NA)
  ) %>%
  select(Country,Scope12, `Scope12:Disclosure1`, Disclosure1)

summary1<- print(xtable(results_df1), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

# Save output in a text file
writeLines(summary1, paste0("7.21_Summary.tex"))

###Table 2
# Combine coefficient and significance into one string
results_df2 <- results_df2 %>%
  mutate(
    Scp12Int = ifelse(!is.na(Coefficient_Scp12Int), paste0(sprintf("%.3f", Coefficient_Scp12Int), Significance_Scp12Int), NA),
    `Scp12Int:Disclosure1` = ifelse(!is.na(`Coefficient_Scp12Int:Disclosure1`), paste0(sprintf("%.3f", `Coefficient_Scp12Int:Disclosure1`), `Significance_Scp12Int:Disclosure1`), NA),
    Disclosure1 = ifelse(!is.na(Coefficient_Disclosure1), paste0(sprintf("%.3f", Coefficient_Disclosure1), Significance_Disclosure1), NA)
  ) %>%
  select(Country,Scp12Int, `Scp12Int:Disclosure1`, Disclosure1)

summary2<- print(xtable(results_df2), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

# Save output in a text file
writeLines(summary2, paste0("7.22_Summary.tex"))

###Table 3
# Combine coefficient and significance into one string
results_df3 <- results_df3 %>%
  mutate(
    Scp12YoY = ifelse(!is.na(Coefficient_Scp12YoY), paste0(sprintf("%.3f", Coefficient_Scp12YoY), Significance_Scp12YoY), NA),
    `Scp12YoY:DisclosureYoY1` = ifelse(!is.na(`Coefficient_Scp12YoY:DisclosureYoY1`), paste0(sprintf("%.3f", `Coefficient_Scp12YoY:DisclosureYoY1`), `Significance_Scp12YoY:DisclosureYoY1`), NA),
    DisclosureYoY1 = ifelse(!is.na(Coefficient_DisclosureYoY1), paste0(sprintf("%.3f", Coefficient_DisclosureYoY1), Significance_DisclosureYoY1), NA)
  ) %>%
  select(Country,Scp12YoY, `Scp12YoY:DisclosureYoY1`, DisclosureYoY1)

summary3<- print(xtable(results_df3), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

# Save output in a text file
writeLines(summary3, paste0("7.23_Summary.tex"))