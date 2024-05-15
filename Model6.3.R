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
  select(DtD, Scope12, Scp12Int, Scp12YoY, Audit, AuditYoY, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree, Year.y.y, Sector_number, Country_number, Sector, Firm)
 

# Create a mapping from Sector_number to Sector
Sector_mapping <- Data %>%
  select(Sector_number, Sector) %>%
  distinct()


#Prepare data for the first model
Data$Audit<-as.factor(Data$Audit)
Data$AuditYoY<-as.factor(Data$AuditYoY)
Data$Year.y.y<-as.factor(Data$Year.y.y)
Data$Country_number<-as.factor(Data$Country_number)

####################
#Models for Level
####################

# Determine the maximum sector number dynamically
max_sector <- max(Data$Sector_number, na.rm = TRUE)

# Initialize an output list
results_df1 <- data.frame(Sector = character(), Variable = character(), Coefficient = numeric(), Significance = character())
results_df2 <- data.frame(Sector = character(), Variable = character(), Coefficient = numeric(), Significance = character())
results_df3 <- data.frame(Sector = character(), Variable = character(), Coefficient = numeric(), Significance = character())
outputs <- list()

# Loop through each sector
for (i in 1:max_sector) {
  SectorData <- filter(Data, Sector_number == i)

  Sector_name <- Sector_mapping$Sector[Sector_mapping$Sector_number == i]
    
  # Model 1
  model1 <- lm(DtD ~ Scope12*Audit + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility  + RiskFree + Year.y.y + Country_number, data = SectorData)
  
  # Model 2
  model2 <- lm(DtD ~ Scp12Int*Audit + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + RiskFree + Year.y.y + Country_number, data = SectorData)
  
  # Model 3
  model3 <- lm(DtD ~ Scp12YoY*AuditYoY + Leverage + log(Size) + Profitability + Liquidity + EquityBuffer + MarketReturn + MarketVolatility + RiskFree + Year.y.y + Country_number, data = SectorData)
  
  # Adjust standard errors for clustering at the firm level
  model1_results <- coeftest(model1, vcov = vcovCL(model1, cluster = ~Firm, data = SectorData))
  model2_results <- coeftest(model2, vcov = vcovCL(model2, cluster = ~Firm, data = SectorData))
  model3_results <- coeftest(model3, vcov = vcovCL(model3, cluster = ~Firm, data = SectorData))
  
  # Capture output of stargazer
  output <- capture.output(
    stargazer(model1, model2, model3, type = "latex",
              coef=list(model1_results[,1], model2_results[,1], model3_results[,1]),
              se=list(model1_results[,2], model2_results[,2], model3_results[,2]),
              t=list(model1_results[,3], model2_results[,3], model3_results[,3]),
              p=list(model1_results[,4], model2_results[,4], model3_results[,4]), 
              title = paste("Regression Models Comparison for",Sector_name),
              omit = c("Year.y.y", "Country_number"), 
              add.lines = list(c("Country fixed-effects", "Yes", "Yes", "Yes"), 
                               c("Time fixed-effects", "Yes", "Yes", "Yes")),
              omit.stat = c("adj.rsq", "ser", "f"), 
              align = TRUE)
  )
  
  # Save output in a text file
  writeLines(output, paste0("6.3_",i,Sector_name, "_models.tex"))
  
  # Optionally store in list to check in R environment
  outputs[[i]] <- output
  
  # Extract coefficients and p-values for each model separately
  tidy1 <- tidy(model1_results)
  tidy2 <- tidy(model2_results)
  tidy3 <- tidy(model3_results)
  
  # Combine and filter the tidy data
  coef_df1 <- tidy1 %>%
    filter(term %in% c("Scope12", "Audit1", "Scope12:Audit1")) %>%
    mutate(Significance = case_when(
      p.value < 0.01 ~ "***",
      p.value < 0.05 ~ "**",
      p.value < 0.10 ~ "*",
      TRUE           ~ ""
    ))
  
  # Combine and filter the tidy data
  coef_df2 <- bind_rows(tidy2)
  coef_df2 <- coef_df2 %>%
    filter(term %in% c("Scp12Int", "Audit1", "Scp12Int:Audit1")) %>%
    mutate(Significance = case_when(
      p.value < 0.01 ~ "***",
      p.value < 0.05  ~ "**",
      p.value < 0.10  ~ "*",
      TRUE            ~ ""
    ))
  
  # Combine and filter the tidy data
  coef_df3 <- bind_rows(tidy3)
  coef_df3 <- coef_df3 %>%
    filter(term %in% c("Scp12YoY", "AuditYoY1","Scp12YoY:AuditYoY1")) %>%
    mutate(Significance = case_when(
      p.value < 0.01 ~ "***",
      p.value < 0.05  ~ "**",
      p.value < 0.10  ~ "*",
      TRUE            ~ ""
    ))
  
  # Append results to the summary dataframe
  for (var in c("Scope12", "Audit1","Scope12:Audit1")) {
    if (any(coef_df1$term == var)) {  # Check if variable exists in the model output
      results_df1 <- rbind(results_df1, data.frame(
        Sector = Sector_name,
        Variable = var,
        Coefficient = coef_df1$estimate[coef_df1$term == var],
        Significance = coef_df1$Significance[coef_df1$term == var]
      ))
    }
  }
  
  # Append results to the summary dataframe
  for (var in c("Scp12Int", "Audit1", "Scp12Int:Audit1")) {
    if (any(coef_df2$term == var)) {  # Check if variable exists in the model output
      results_df2 <- rbind(results_df2, data.frame(
        Sector = Sector_name,
        Variable = var,
        Coefficient = coef_df2$estimate[coef_df2$term == var],
        Significance = coef_df2$Significance[coef_df2$term == var]
      ))
    }
  }
  
  # Append results to the summary dataframe
  for (var in c("Scp12YoY", "AuditYoY1","Scp12YoY:AuditYoY1")) {
    if (any(coef_df3$term == var)) {  # Check if variable exists in the model output
      results_df3 <- rbind(results_df3, data.frame(
        Sector = Sector_name,
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
    `Scope12:Audit1` = ifelse(!is.na(`Coefficient_Scope12:Audit1`), paste0(sprintf("%.3f", `Coefficient_Scope12:Audit1`), `Significance_Scope12:Audit1`), NA),
    Audit1 = ifelse(!is.na(Coefficient_Audit1), paste0(sprintf("%.3f", Coefficient_Audit1), Significance_Audit1), NA)
  ) %>%
  select(Sector,Scope12, `Scope12:Audit1`, Audit1)

summary1<- print(xtable(results_df1), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

# Save output in a text file
writeLines(summary1, paste0("6.31_Summary.tex"))

###Table 2
# Combine coefficient and significance into one string
results_df2 <- results_df2 %>%
  mutate(
    Scp12Int = ifelse(!is.na(Coefficient_Scp12Int), paste0(sprintf("%.3f", Coefficient_Scp12Int), Significance_Scp12Int), NA),
    `Scp12Int:Audit1` = ifelse(!is.na(`Coefficient_Scp12Int:Audit1`), paste0(sprintf("%.3f", `Coefficient_Scp12Int:Audit1`), `Significance_Scp12Int:Audit1`), NA),
    Audit1 = ifelse(!is.na(Coefficient_Audit1), paste0(sprintf("%.3f", Coefficient_Audit1), Significance_Audit1), NA)
  ) %>%
  select(Sector,Scp12Int, `Scp12Int:Audit1`, Audit1)

summary2<- print(xtable(results_df2), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

# Save output in a text file
writeLines(summary2, paste0("6.32_Summary.tex"))

###Table 3
# Combine coefficient and significance into one string
results_df3 <- results_df3 %>%
  mutate(
    Scp12YoY = ifelse(!is.na(Coefficient_Scp12YoY), paste0(sprintf("%.3f", Coefficient_Scp12YoY), Significance_Scp12YoY), NA),
    `Scp12YoY:AuditYoY1` = ifelse(!is.na(`Coefficient_Scp12YoY:AuditYoY1`), paste0(sprintf("%.3f", `Coefficient_Scp12YoY:AuditYoY1`), `Significance_Scp12YoY:AuditYoY1`), NA),
    AuditYoY1 = ifelse(!is.na(Coefficient_AuditYoY1), paste0(sprintf("%.3f", Coefficient_AuditYoY1), Significance_AuditYoY1), NA)
  ) %>%
  select(Sector,Scp12YoY, `Scp12YoY:AuditYoY1`, AuditYoY1)

summary3<- print(xtable(results_df3), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

# Save output in a text file
writeLines(summary3, paste0("6.33_Summary.tex"))