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

#step 1+2 only necessary the first time
font_path <- "C:/Users/mathi/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code"
font_import(paths = font_path, prompt = FALSE)

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
#Load Ratings
load("SP_Rating.Rdata")

#Data including scope 3
Reg_dataInclscp3 <- DtD.results%>%
  full_join(Scp1, by=c("Firm","Date"))

number_of_firms <- length(unique(Data_Wins$Firm))

# Print the result
print(number_of_firms)

##Variables
#Date
#Firm
#DtD
#Scope1
#Scope2
#Scope3
#Scope12
#Scope1_2
#Scp1Int
#Scp2Int
#Scp3Int
#Scp12Int
#Scp1_2Int
#Scp12YoY
#Scp1_2YoY
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
#Correlation of scopes
##################################

# Filter on firms with only 1 observations
filtered_data1 <- Reg_Data %>%
  group_by(Firm) %>%
  filter(n() > 1) %>%
  ungroup()

# Compute correlation matrix for each firm
list_of_cor_matrices1 <- filtered_data1 %>%
  group_by(Firm) %>%
  group_split() %>%
  map(~ {
    data_subset <- .x %>% select(Scope1, Scope2)
    # Handle the case where any variable has zero variance or NA variance
    if (any(is.na(sapply(data_subset, var))) || any(sapply(data_subset, var) == 0)) {
      return(NA)  # Return NA matrix explicitly
    }
    # Compute correlation
    cor(data_subset, use = "complete.obs")
  })

# Filter out matrices that are entirely NA or where calculation was not possible
list_of_cor_matrices1 <- list_of_cor_matrices1[!sapply(list_of_cor_matrices1, function(x) all(is.na(x), na.rm = TRUE))]

# Average the correlation matrices
mean_cor_matrix1 <- Reduce("+", list_of_cor_matrices1) / length(list_of_cor_matrices1)

##################################
#Data selection and final sample
##################################
# Define row names
row_names <- c("Initial sample", 
               "- Financial firms", 
               "- less than 1 year of observations", 
               "- Missing values for DtD", 
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

##Missing values for control variables
load("DtD_results.Rdata")
Data_Selection[4, 3] <- length(unique(DtD.results$Firm))

##Missing values for Scope 1 & 2
Data_Selection[5, 3] <- length(unique(Data_Wins$Firm))

##Final sample
Data_Selection[6, 3] <- length(unique(Data_Wins$Firm))

#Fill out firms excluded
for (i in 2:(nrow(Data_Selection) - 1)) {
  Data_Selection[i, 2] <- Data_Selection[i - 1, 3] - Data_Selection[i, 3]
}

# Print the matrix
print(Data_Selection)

##################################
#Data composition, by firm
##################################
# Select relevant columns and filter for distinct records
Data_comp <- Data_Wins %>%
  select(Year.y.y, Firm, Country, Sector) %>%
  distinct()

# Summarize by Year
year_summary <- Data_comp %>%
  group_by(Year.y.y) %>%
  summarize(Obs_Y = n(), .groups = 'drop')

# Summarize by Country
country_summary <- Data_comp %>%
  group_by(Country) %>%
  summarize(Obs_C = n(), .groups = 'drop')

# Summarize by Sector
sector_summary <- Data_comp %>%
  group_by(Sector) %>%
  summarize(Obs_S = n(), .groups = 'drop')

# Summarize by Sector
sector_summary <- Data_comp %>%
  group_by(Sector) %>%
  summarize(Obs_S = n(), .groups = 'drop')

# Add an index to each summary for merging
year_summary$Index <- seq_along(year_summary$Year.y.y)
country_summary$Index <- seq_along(country_summary$Country)
sector_summary$Index <- seq_along(sector_summary$Sector)

# Merge the summaries using full_join to ensure all data is kept
combined_df <- full_join(year_summary, country_summary, by = "Index", suffix = c(".year", ".country"))
combined_df <- full_join(combined_df, sector_summary, by = "Index")

# Cleanup: Remove the Index column and convert all NAs in Year.y.y, Country, and Sector
combined_df$Index <- NULL
combined_df$Year.y.y <- as.character(combined_df$Year.y.y)
combined_df$Country <- as.character(combined_df$Country)
combined_df$Sector <- as.character(combined_df$Sector)


final_combined_df <- bind_rows(combined_df)

# Calculate the totals for each 'Obs' column
total_by_year <- sum(combined_df$Obs_Y, na.rm = TRUE)
total_by_country <- sum(combined_df$Obs_C, na.rm = TRUE)
total_by_sector <- sum(combined_df$Obs_S, na.rm = TRUE)

# Create the first summary row with totals per category
summary_row1 <- data.frame(
  Year.y.y = "Obs.",
  Obs_Y = total_by_year,
  Country = "Obs.",
  Obs_C = total_by_country,
  Sector = "Obs.",
  Obs_S = total_by_sector
)

Firms <- Data_Wins %>%
  distinct(Firm) %>%
  group_by(Firm) %>%
  summarize(Firm = n(), .groups = 'drop')

# Create the second summary row with grand total
total_firms <- nrow(Firms)
summary_row2 <- data.frame(
  Year.y.y = "Firms",
  Obs_Y = total_firms,
  Country = "Firms",
  Obs_C = total_firms,
  Sector = "Firms",
  Obs_S = total_firms
)

# Append both summary rows to the combined data frame
Data_composition <- bind_rows(combined_df, summary_row1, summary_row2)

##################################
#Correlation of climate variables
##################################

# Filter on firms with only 1 observations
filtered_data2 <- Reg_Data %>%
  group_by(Firm) %>%
  filter(n() > 1) %>%
  ungroup()

# Compute correlation matrix for each firm
list_of_cor_matrices2 <- filtered_data2 %>%
  group_by(Firm) %>%
  group_split() %>%
  map(~ {
    data_subset <- .x %>% select(Scope12, Scp12Int, Scp12YoY)
    # Handle the case where any variable has zero variance or NA variance
    if (any(is.na(sapply(data_subset, var))) || any(sapply(data_subset, var) == 0)) {
      return(NA)  # Return NA matrix explicitly
    }
    # Compute correlation
    cor(data_subset, use = "complete.obs")
  })

# Filter out matrices that are entirely NA or where calculation was not possible
list_of_cor_matrices2 <- list_of_cor_matrices2[!sapply(list_of_cor_matrices2, function(x) all(is.na(x), na.rm = TRUE))]

# Average the correlation matrices
mean_cor_matrix2 <- Reduce("+", list_of_cor_matrices2) / length(list_of_cor_matrices2)

###############################
#Correlation before winsorizing
###############################

# Filter on firms with only 1 observations
filtered_data3 <- Reg_Data %>%
  group_by(Firm) %>%
  filter(n() > 1) %>%
  ungroup()

# Compute correlation matrix for each firm
list_of_cor_matrices3 <- filtered_data3 %>%
  group_by(Firm) %>%
  group_split() %>%
  map(~ {
    data_subset <- .x %>% select(DtD, Scope12, Scp12Int, Scp12YoY, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree)
    # Handle the case where any variable has zero variance or NA variance
    if (any(is.na(sapply(data_subset, var))) || any(sapply(data_subset, var) == 0)) {
      return(NA)  # Return NA matrix explicitly
    }
    # Compute correlation
    cor(data_subset, use = "complete.obs")
  })

# Filter out matrices that are entirely NA or where calculation was not possible
list_of_cor_matrices3 <- list_of_cor_matrices3[!sapply(list_of_cor_matrices3, function(x) all(is.na(x), na.rm = TRUE))]

# Average the correlation matrices
mean_cor_matrix3 <- Reduce("+", list_of_cor_matrices3) / length(list_of_cor_matrices3)


###############################
#Correlation After winsorizing
###############################

# Filter on firms with only 1 observations
filtered_data4 <- Data_Wins %>%
  group_by(Firm) %>%
  filter(n() > 1) %>%
  ungroup()

# Compute correlation matrix for each firm
list_of_cor_matrices4 <- filtered_data4 %>%
  group_by(Firm) %>%
  group_split() %>%
  map(~ {
    data_subset <- .x %>% select(DtD, Scope12, Scp12Int, Scp12YoY, Size, Leverage, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree)
    # Handle the case where any variable has zero variance or NA variance
    if (any(is.na(sapply(data_subset, var))) || any(sapply(data_subset, var) == 0)) {
      return(NA)  # Return NA matrix explicitly
    }
    # Compute correlation
    cor(data_subset, use = "complete.obs")
  })

# Filter out matrices that are entirely NA or where calculation was not possible
list_of_cor_matrices4 <- list_of_cor_matrices4[!sapply(list_of_cor_matrices4, function(x) all(is.na(x), na.rm = TRUE))]

# Average the correlation matrices
mean_cor_matrix4 <- Reduce("+", list_of_cor_matrices4) / length(list_of_cor_matrices4)

#############################
#Descriptive statistic table
#############################

#Table of data not winsorized
Summary_table1_data<-Reg_Data%>%
  select(DtD, Scope1, Scope2, Scope12R, Scope12, Scp12Int, Scp12YoY, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree)

#Set data to procent
Summary_table1_data$Scp12YoY<-Summary_table1_data$Scp12YoY*100
Summary_table1_data$Profitability<-Summary_table1_data$Profitability*100
Summary_table1_data$MarketReturn<-Summary_table1_data$MarketReturn*100
Summary_table1_data$Inflation<-Summary_table1_data$Inflation*100

#Rename name variables for tables
new_names <- c(
  "DtD" = "Distance to Default",
  "Scope1" = "Reported Scope 1",
  "Scope2" = "Reported Scope 2",
  "Scope12R" = "Reported Scope 1&2",
  "Scope12" = "Scope 1&2 level",
  "Scp12Int" = "Scope 1&2 Intensity",
  "Scp12YoY" = "Scope 1&2 level YoY change(%)",
  "Leverage" = "Leverage",
  "Size" = "Size",
  "Profitability" = "Profitability(%)",
  "Liquidity" = "Liquidity",
  "EquityBuffer" = "Equity Buffer",
  "MarketReturn" = "Market Return(%)",
  "MarketVolatility" = "Market Volatility(%)",
  "Inflation" = "Inflation(%)",
  "RiskFree" = "Risk Free(%)"
)

# Creating summary statistics table
Summary_table1 <- Summary_table1_data %>%
  summarise(across(everything(), list(
    Mean = ~mean(., na.rm = TRUE),
    Median = ~median(., na.rm = TRUE),
    Min = ~min(., na.rm = TRUE),
    Max = ~max(., na.rm = TRUE),
    StdDev = ~sd(., na.rm = TRUE),
    CoefVar = ~sd(., na.rm = TRUE) / mean(., na.rm = TRUE) * 100  # Calculate Coefficient of Variation
  ))) %>%
  pivot_longer(
    cols = everything(),
    names_to = c("Variable", ".value"),
    names_sep = "_"
  )

Summary_table1 <- Summary_table1 %>%
  mutate(Variable = new_names[Variable])

#Table of winsorized data
Summary_table2_data<-Data_Wins%>%
  select(DtD, Scope1, Scope2, Scope12R, Scope12, Scp12Int, Scp12YoY, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree)

#Set data to procent
Summary_table2_data$Scp12YoY<-Summary_table2_data$Scp12YoY*100
Summary_table2_data$Profitability<-Summary_table2_data$Profitability*100
Summary_table2_data$MarketReturn<-Summary_table2_data$MarketReturn*100
Summary_table2_data$Inflation<-Summary_table2_data$Inflation*100

Summary_table2 <- Summary_table2_data %>%
  summarise(across(everything(), list(
    Mean = ~mean(., na.rm = TRUE),
    Median = ~median(., na.rm = TRUE),
    Min = ~min(., na.rm = TRUE),
    Max = ~max(., na.rm = TRUE),
    StdDev = ~sd(., na.rm = TRUE),
    CoefVar = ~sd(., na.rm = TRUE) / mean(., na.rm = TRUE) * 100  # Calculate Coefficient of Variation
  ))) %>%
  pivot_longer(
    cols = everything(),
    names_to = c("Variable", ".value"),
    names_sep = "_"
  )

Summary_table2 <- Summary_table2 %>%
  mutate(Variable = new_names[Variable])

rm(Summary_table1_data, Summary_table2_data)

#Get Latex codes
latex_code1 <- print(xtable(Data_Selection), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)
latex_code2 <- print(xtable(Data_composition), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)
latex_code3 <- print(xtable(Summary_table1), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)
latex_code4 <- print(xtable(Summary_table2), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)
latex_code5 <- print(xtable(mean_cor_matrix1), include.rownames = TRUE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)
latex_code6 <- print(xtable(mean_cor_matrix2), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

#########################
#Plot data
#########################

Sector_Mean_Scope12 <- Data_Wins %>%
  group_by(Sector) %>%
  summarise(Mean_Scope12 = mean(Scope12, na.rm = TRUE))%>%
  arrange(desc(Mean_Scope12)) 

Scope_sector <- ggplot(Sector_Mean_Scope12, aes(x = reorder(Sector, -Mean_Scope12), y = Mean_Scope12)) +
  geom_bar(stat = "identity", fill = "darkblue") +
  labs(
    x = "Sector",
    y = "Average Scope 1 & 2",
    title = "",
    color = ""
  ) +
  theme_minimal() +
  theme(
    text = element_text(family = "CMU Serif"),
    legend.position = "right",
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.text.x = element_text(angle = 45, hjust = 1),  # Skew x-axis labels
    axis.text.y = element_blank(),  # Removes Y-axis text labels
    axis.ticks.y = element_blank(),  # Removes Y-axis ticks
    axis.line = element_line(colour = "black"),  # Ensures axis lines are visible
    axis.title.x = element_text(margin = margin(t = 10, unit = "pt")),  # Increase gap between x-axis title and labels
    plot.margin = margin(l = 50, unit = "pt")  # Adjusts plot margins
  )

print(Scope_sector)

ggsave("Scope by sector.png", plot = Scope_sector, width = 6, height = 3, dpi = 300)
  
  
#########################
#Misc
#########################
#if (!require(DescTools)) install.packages("DescTools")
#library(DescTools)

#Assets_win<- Assets %>%
 # group_by(Year) %>%
  #mutate(Assets = Winsorize(Assets, probs = c(0.01, 0.99)))

#summary_Assets <- Assets %>%
  #group_by(Year) %>%
  #summarise(min_value = min(Assets)/10^7,
            #max_value = max(Assets)/10^7,
           # median_value = median(Assets)/10^7,
            #Mean_value=mean(Assets)/10^7)

#summary_Assets_win <- Assets_win %>%
 # group_by(Year) %>%
  #summarise(min_value = min(Assets)/10^7,
    #        max_value = max(Assets)/10^7,
   #         median_value = median(Assets)/10^7,
     #       Mean_value=mean(Assets)/10^7)


#ggplot(summary_Assets, aes(x = Year)) +
 # geom_segment(aes(x = Year, xend = Year, y = min_value, yend = max_value), color = "gray") +  # Vertical lines
  #geom_point(aes(y = median_value), color = "blue", size = 3) +  # Median points
  #geom_point(aes(y = Mean_value), color = "green", size = 3) +  # Mean points
  #labs(title = "Yearly Value Range with Median",
  #     x = "Year",
   #    y = "Scope1") +
  #theme_minimal()

#ggplot(summary_Assets_win, aes(x = Year)) +
 # geom_segment(aes(x = Year, xend = Year, y = min_value, yend = max_value), color = "gray") +  # Vertical lines
  #geom_point(aes(y = median_value), color = "blue", size = 3) +  # Median points
  #geom_point(aes(y = Mean_value), color = "green", size = 3) +  # Mean points
  #labs(title = "Yearly Value Range with Median",
  #     x = "Year",
  #     y = "Scope1") +
  #theme_minimal()