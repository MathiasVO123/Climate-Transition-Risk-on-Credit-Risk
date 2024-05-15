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
  select(DtD, Scope12, Scp12Int, Scp12YoY, SBTiTarget, SBTiTargetYoY, Leverage, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , Inflation , RiskFree, Year.y.y, Sector_number, Country_number, Firm)

#Prepare data for the first model
Data$SBTiTarget<-as.factor(Data$SBTiTarget)
Data$Year.y.y<-as.factor(Data$Year.y.y)
Data$Sector_number<-as.factor(Data$Sector_number)
Data$Country_number<-as.factor(Data$Country_number)

Data$sector_country <- interaction(Data$Sector_number, Data$Country_number)
Data$sector_country <- as.factor(Data$sector_country)

####################
#Models
####################

#Model 1
model1 <- lm(DtD ~ Scope12*SBTiTarget + Leverage + log(Size)+ Profitability+ Liquidity + EquityBuffer+ MarketReturn + MarketVolatility + RiskFree + Year.y.y + sector_country, data = Data)
summary(model1)

#Model 2
model2 <- lm(DtD ~ Scp12Int*SBTiTarget + Leverage + log(Size)+ Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = Data)
summary(model2)

#Model 3
model3 <- lm(DtD ~ Scp12YoY*SBTiTargetYoY + Leverage + log(Size)+ Profitability+ Liquidity+ EquityBuffer+ MarketReturn + MarketVolatility + RiskFree  + Year.y.y + sector_country, data = Data)
summary(model3)

# Adjust standard errors for clustering at the firm level
model1_results <- coeftest(model1, vcov = vcovCL(model1, cluster = ~Firm, data = Data))
model2_results <- coeftest(model2, vcov = vcovCL(model2, cluster = ~Firm, data = Data))
model3_results <- coeftest(model3, vcov = vcovCL(model3, cluster = ~Firm, data = Data))

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

#########################
#Assumptions Linearity
#########################
#Model 1
residuals_df1 <- data.frame(
  Fitted = fitted(model1),
  Residuals1 = residuals(model1)
)

# Create the residual plot using ggplot2
residual_plot1 <- ggplot(residuals_df1, aes(x = Fitted, y = Residuals1)) +
  geom_point() +  # Plot the points
  geom_hline(yintercept = 0, color = "red", linetype = "dashed") +  # Add a horizontal line at y = 0
  labs(x = "Fitted Values",y = "Residuals",title = "",color="") +
  theme_minimal() +
  theme(
    text = element_text(family = "CMU Serif"),
    legend.position="right",
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.text.y = element_blank(),  # Removes Y-axis text labels
    axis.ticks.y = element_blank(),  # Removes Y-axis ticks
    axis.line = element_line(colour = "black"),  # Ensures axis lines are visible
    plot.margin = margin(l = 50, unit = "pt"))

ggsave("residual_plot14.png", plot = residual_plot1, width = 6, height = 3, dpi = 300)

#Model 2
residuals_df2 <- data.frame(
  Fitted = fitted(model2),
  Residuals2 = residuals(model2)
)

# Create the residual plot using ggplot2
residual_plot2 <- ggplot(residuals_df2, aes(x = Fitted, y = Residuals2)) +
  geom_point() +  # Plot the points
  geom_hline(yintercept = 0, color = "red", linetype = "dashed") +  # Add a horizontal line at y = 0
  labs(x = "Fitted Values",y = "Residuals",title = "",color="") +
  theme_minimal() +
  theme(
    text = element_text(family = "CMU Serif"),
    legend.position="right",
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.line = element_line(colour = "black"),  # Ensures axis lines are visible
    plot.margin = margin(l = 50, unit = "pt"))

ggsave("residual_plot24.png", plot = residual_plot2, width = 6, height = 3, dpi = 300)

#Model 3
residuals_df3 <- data.frame(
  Fitted = fitted(model3),
  Residuals3 = residuals(model3)
)

# Create the residual plot using ggplot3
residual_plot3 <- ggplot(residuals_df3, aes(x = Fitted, y = Residuals3)) +
  geom_point() +  # Plot the points
  geom_hline(yintercept = 0, color = "red", linetype = "dashed") +  # Add a horizontal line at y = 0
  labs(x = "Fitted Values",y = "Residuals",title = "",color="") +
  theme_minimal() +
  theme(
    text = element_text(family = "CMU Serif"),
    legend.position="right",
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.line = element_line(colour = "black"),  # Ensures axis lines are visible
    plot.margin = margin(l = 50, unit = "pt"))

ggsave("residual_plot34.png", plot = residual_plot3, width = 6, height = 3, dpi = 300)

#########################
#Assumptions Rank
#########################
##Model 1
#Rank
DataRank1<-Data%>%
  select(Scope12, SBTiTarget, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , RiskFree)%>%
  na.omit()

DataRank1 <- as.matrix(DataRank1)
original_colnames <- colnames(DataRank1)
DataRank1 <- matrix(as.numeric(DataRank1), nrow = nrow(DataRank1), ncol = ncol(DataRank1))
colnames(DataRank1) <- original_colnames
DataRank1 <- DataRank1[1:11, ]
DataRank1<-rref(DataRank1)
Rank1 <- xtable(DataRank1, align = rep("r", ncol(DataRank1) + 1), digits = 0)  # `digits=0` if it's an integer matrix

#VIF
VIF1 <- update(model1, . ~ . - Year.y.y - sector_country)
summary(VIF1)
vif_values1 <- vif(VIF1)
vIF1 <- data.frame(
  Variable = names(vif_values1),
  VIF = unname(vif_values1)
)
VIF1 <- print(xtable(vIF1), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)


##Model 2
DataRank2<-Data%>%
  select(Scp12Int, SBTiTarget, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , RiskFree)%>%
  na.omit()

DataRank2 <- as.matrix(DataRank2)
original_colnames <- colnames(DataRank2)
DataRank2 <- matrix(as.numeric(DataRank2), nrow = nrow(DataRank2), ncol = ncol(DataRank2))
colnames(DataRank2) <- original_colnames
DataRank2 <- DataRank2[1:11, ]
DataRank2<-rref(DataRank2)
Rank2 <- xtable(DataRank2, align = rep("r", ncol(DataRank2) + 1), digits = 0)  # `digits=0` if it's an integer matrix


#VIF
VIF2 <- update(model2, . ~ . - Year.y.y - sector_country)
summary(VIF2)
vif_values2 <- vif(VIF2)
vIF2 <- data.frame(
  Variable = names(vif_values2),
  VIF = unname(vif_values2)
)
VIF2 <- print(xtable(vIF2), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

##Model 3
DataRank3<-Data%>%
  select(Scp12YoY, SBTiTarget, Size, Profitability, Liquidity, EquityBuffer, MarketReturn , MarketVolatility , RiskFree)%>%
  na.omit()

DataRank3 <- as.matrix(DataRank3)
original_colnames <- colnames(DataRank3)
DataRank3 <- matrix(as.numeric(DataRank3), nrow = nrow(DataRank3), ncol = ncol(DataRank3))
colnames(DataRank3) <- original_colnames
DataRank3 <- DataRank3[1:11, ]
DataRank3<-rref(DataRank3)
Rank3 <- xtable(DataRank3, align = rep("r", ncol(DataRank3) + 1), digits = 0)  # `digits=0` if it's an integer matrix


#VIF
VIF3 <- update(model3, . ~ . - Year.y.y - sector_country)
summary(VIF3)
vif_values3 <- vif(VIF3)
vIF3 <- data.frame(
  Variable = names(vif_values3),
  VIF = unname(vif_values3)
)
VIF3 <- print(xtable(vIF3), include.rownames = FALSE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

#########################
#Assumptions Exogeneity 
#########################

#Model 1
# Calculate residuals from the original model
residuals1 <- residuals(model1)

# Regress residuals on original independent variables
original1 <- formula(model1)
dependent_variable <- as.character(original1[[2]])  # Extract dependent variable
modified_formula <- update(original1, paste(". ~ . -", dependent_variable))

Res_model1 <- lm(update(modified_formula, residuals1 ~ .), data = Data)
Res_summary1 <- summary(Res_model1)
Rescoe1 <- as.data.frame(Res_summary1$coefficients)

Rescoe1 <- Rescoe1[!grepl("^Year\\.y\\.y|^sector_country", rownames(Rescoe1)), ]

Rescoe1 <- Rescoe1[2:20, -2]

Res1 <- print(xtable(Rescoe1), include.rownames = TRUE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

#Model 2
# Calculate residuals from the original model
residuals2 <- residuals(model2)

# Regress residuals on original independent variables
original2 <- formula(model2)
dependent_variable <- as.character(original2[[2]])  # Extract dependent variable
modified_formula <- update(original2, paste(". ~ . -", dependent_variable))

Res_model2 <- lm(update(modified_formula, residuals2 ~ .), data = Data)
Res_summary2 <- summary(Res_model2)
Rescoe2 <- as.data.frame(Res_summary2$coefficients)

Rescoe2 <- Rescoe2[!grepl("^Year\\.y\\.y|^sector_country", rownames(Rescoe2)), ]

Rescoe2 <- Rescoe2[2:20, -2]

Res2 <- print(xtable(Rescoe2), include.rownames = TRUE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

#Model 3
# Calculate residuals from the original model
residuals3 <- residuals(model2)

# Regress residuals on original independent variables
original3 <- formula(model3)
dependent_variable <- as.character(original3[[2]])  # Extract dependent variable
modified_formula <- update(original3, paste(". ~ . -", dependent_variable))

Res_model3 <- lm(update(modified_formula, residuals3 ~ .), data = Data)
Res_summary3 <- summary(Res_model3)
Rescoe3 <- as.data.frame(Res_summary3$coefficients)

Rescoe3 <- Rescoe3[!grepl("^Year\\.y\\.y|^sector_country", rownames(Rescoe3)), ]

Rescoe3 <- Rescoe3[2:20, -2]

Res3 <- print(xtable(Rescoe3), include.rownames = TRUE, include.colnames = TRUE, floating = FALSE, hline.after = NULL, print.results = TRUE)

##############################
#Assumptions Homoscedasticity 
##############################

#Model 1
dw_result1 <- dwtest(model1)

# Print the results
print(dw_result1)

#Model 2
dw_result2 <- dwtest(model2)

# Print the results
print(dw_result2)

#Model 3
dw_result3 <- dwtest(model3)

# Print the results
print(dw_result3)

#########################
#Assumptions Normality
#########################

#Model 1
# Create data for Q-Q plot
qq_data1 <- qqnorm(residuals1, plot.it = FALSE)
qq_df1 <- data.frame(Theoretical = qq_data1$x, Sample = qq_data1$y)

#Normal probability plot
qq_plot1 <- ggplot(qq_df1, aes(sample = Sample)) +
  stat_qq() +
  stat_qq_line(colour = "red") +
  labs(title = "", x = "Normal quantiles", y = "Sample Quantiles") +
  theme_minimal() +
  theme(text = element_text(family = "CMU Serif"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_line(colour = "black"))

#Histogram
hist_plot1 <- ggplot(residuals_df1, aes(x = Residuals1)) +
  geom_histogram(bins = 50, fill = "grey", color = "black") +
  labs(title = "", x = "Residuals", y = "Frequency") +
  theme_minimal() +
  theme(text = element_text(family = "CMU Serif"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_line(colour = "black"))

qq_plot1 <- qq_plot1 +
  theme(plot.margin = margin(t = 10, r = 20, b = 10, l = 10, unit = "pt"))  # Increase right margin

hist_plot1 <- hist_plot1 +
  theme(plot.margin = margin(t = 10, r = 10, b = 10, l = 20, unit = "pt"))  # Increase left margin

# Combine the plots with adjusted margins
combined_plot1 <- grid.arrange(qq_plot1, hist_plot1, ncol = 2)

# Save the combined plot with appropriate dimensions
ggsave("Normal14.png", combined_plot1, width = 12, height = 6)

#Model 2
# Create data for Q-Q plot
qq_data2 <- qqnorm(residuals2, plot.it = FALSE)
qq_df2 <- data.frame(Theoretical = qq_data2$x, Sample = qq_data2$y)

#Normal probability plot
qq_plot2 <- ggplot(qq_df2, aes(sample = Sample)) +
  stat_qq() +
  stat_qq_line(colour = "red") +
  labs(title = "", x = "Normal quantiles", y = "Sample Quantiles") +
  theme_minimal() +
  theme(text = element_text(family = "CMU Serif"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_line(colour = "black"))

#Histogram
hist_plot2 <- ggplot(residuals_df2, aes(x = Residuals2)) +
  geom_histogram(bins = 50, fill = "grey", color = "black") +
  labs(title = "", x = "Residuals", y = "Frequency") +
  theme_minimal() +
  theme(text = element_text(family = "CMU Serif"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_line(colour = "black"))

qq_plot2 <- qq_plot2 +
  theme(plot.margin = margin(t = 10, r = 20, b = 10, l = 10, unit = "pt"))  # Increase right margin

hist_plot2 <- hist_plot2 +
  theme(plot.margin = margin(t = 10, r = 10, b = 10, l = 20, unit = "pt"))  # Increase left margin

# Combine the plots with adjusted margins
combined_plot2 <- grid.arrange(qq_plot2, hist_plot2, ncol = 2)

# Save the combined plot with appropriate dimensions
ggsave("Normal24.png", combined_plot2, width = 12, height = 6)

#Model 3
# Create data for Q-Q plot
qq_data3 <- qqnorm(residuals3, plot.it = FALSE)
qq_df3 <- data.frame(Theoretical = qq_data3$x, Sample = qq_data3$y)

#Normal probability plot
qq_plot3 <- ggplot(qq_df3, aes(sample = Sample)) +
  stat_qq() +
  stat_qq_line(colour = "red") +
  labs(title = "", x = "Normal quantiles", y = "Sample Quantiles") +
  theme_minimal() +
  theme(text = element_text(family = "CMU Serif"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_line(colour = "black"))

#Histogram
hist_plot3 <- ggplot(residuals_df3, aes(x = Residuals3)) +
  geom_histogram(bins = 50, fill = "grey", color = "black") +
  labs(title = "", x = "Residuals", y = "Frequency") +
  theme_minimal() +
  theme(text = element_text(family = "CMU Serif"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_line(colour = "black"))

qq_plot3 <- qq_plot3 +
  theme(plot.margin = margin(t = 10, r = 20, b = 10, l = 10, unit = "pt"))  # Increase right margin

hist_plot3 <- hist_plot3 +
  theme(plot.margin = margin(t = 10, r = 10, b = 10, l = 20, unit = "pt"))  # Increase left margin

# Combine the plots with adjusted margins
combined_plot3 <- grid.arrange(qq_plot3, hist_plot3, ncol = 2)

# Save the combined plot with appropriate dimensions
ggsave("Normal34.png", combined_plot3, width = 12, height = 6)

