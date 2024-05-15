library(ggplot2)
library(extrafont)

#removing all previous variables 
rm(list=ls())

#setting WD to code folder  
setwd("C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/3. R code")
list.files()

#Downloading the correct font
#step 1+2 only necessary the first time
#font_path <- "C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/3. R code/font.ttf"
#font_import(paths = font_path, prompt = FALSE)

loadfonts()

#####################################
# Merton Model
#####################################

# Parameters
T <- 10 # Random later point
N <- 1000 # Number of steps
dt <- T/N # Time increment
t <- seq(0, T, by = dt) # Time sequence
initial_value <- 100 # Initial asset value
mu <- 0.05 # Drift
sigma <- 0.20 # Volatility

##############################################################################
# Simulating Brownian Motions
#W1 <- c(0, cumsum(rnorm(N, mean = 0, sd = sqrt(dt))))
#W2 <- c(0, cumsum(rnorm(N, mean = 0, sd = sqrt(dt))))
#S1 <- initial_value * exp((mu - 0.5 * sigma^2) * t + sigma * W1)
#S2 <- initial_value * exp(((mu-0.01) - 0.5 * sigma^2) * t + sigma * W2)

# Data frame for ggplot
#Simulated_BM <- data.frame(t = rep(t, 2), 
                   #S = c(S1, S2), 
                   #Firm = factor(rep(1:2, each = N + 1)))

#save(Simulated_BM,file="Simulated_BM.Rdata")
###############################################################################

#loading the simulated data producing nice graphs 

load("Simulated_BM.Rdata")

# Debt level
debt_level <- 80 # Random debt level



#Make plot 
Merton_plot<-ggplot(Simulated_BM, aes(x = t, y = S, color = Firm)) +
  geom_line(size=0.6) +
  scale_x_continuous(
    limits=c(0,T+1),
    breaks = c(0, T),  # Set breaks at 0 and T
    labels = c("t=0", paste0("t=T")),  # Custom labels for 0 and T
    expand = c(0, 0)  # No expansion
  ) +
  geom_vline(xintercept = T, linetype = "dashed") +
  geom_segment(aes(x = 0, y = debt_level, xend = 10, yend = debt_level),linetype = "dashed",col="black") +
  annotate("text", x = T+0.15, y = debt_level, label = "D", hjust = 0.1) +
  labs(x = "Time", y = "Asset Value", title = "",color="") +
  scale_color_manual(values = c("lightblue", "darkblue"), labels = c("Firm A", "Firm B"))+
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


ggsave("Merton Model.png", plot = Merton_plot, width = 6, height = 3, dpi = 300)

##############################################################################
#Creating Merton Pay-Off graphs 
##############################################################################

FV_debt<-225
asset_value<-seq(0,500,by=10)
equity<-pmax(asset_value-FV_debt,0)
debt<-debt_level-pmax(FV_debt-asset_value,0)
payoff_data<-data.frame(asset_value=asset_value, equity_payoff=equity, debt_payoff=debt)

equity_payoff<-ggplot(payoff_data, aes(x = asset_value, y = equity_payoff)) +
  geom_line(size = 0.8) +
  labs(title = "", x = "Firm Asset Value", y = "Pay-off to Equity owners") +
  theme_minimal() +
  theme(
    text = element_text(family = "CMU Serif"),
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.text.y = element_blank(),  # Removes Y-axis text labels
    axis.ticks.y = element_blank(),  # Removes Y-axis ticks
    axis.line = element_line(colour = "black")  # Ensures axis lines are visible
  ) +
  scale_x_continuous(breaks = c(FV_debt), labels = c("D"),expand = c(0, 0))+
  scale_y_continuous(expand=expansion(mult = c(0, 0.5)))

debt_payoff<-ggplot(payoff_data, aes(x = asset_value, y = debt_payoff)) +
  geom_line(size = 0.8) +
  labs(title = "", x = "Firm Asset Value", y = "Pay-off to Debt holders") +
  theme_minimal() +
  theme(
    text = element_text(family = "CMU Serif"),
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.text.y = element_blank(),  # Removes Y-axis text labels
    axis.ticks.y = element_blank(),  # Removes Y-axis ticks
    axis.line = element_line(colour = "black")  # Ensures axis lines are visible
  ) +
  scale_x_continuous(breaks = c(FV_debt), labels = c("D"),expand = c(0, 0))+
  scale_y_continuous(expand = expansion(mult = c(0, 1)))

plot(debt_payoff)

ggsave("Equity Payoff.png", plot = equity_payoff, width = 3.5, height = 3, units = "in")
ggsave("Debt Payoff.png", plot = debt_payoff, width = 3.5, height = 3, units = "in")