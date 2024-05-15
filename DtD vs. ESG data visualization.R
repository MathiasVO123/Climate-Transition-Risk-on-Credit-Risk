library(xtable)
library(readxl)
library(xts)
library(tidyr)
library(dplyr)
library(zoo)
library(lubridate)
library(ggplot2)
library(ggplot2)
library(extrafont)
library(gridExtra)

#loading font
loadfonts()
#removing all previous variables 
rm(list=ls())

##########################
#LOAD DATA AND FUNCTIONS
##########################

#setting WD to code folder 
setwd("C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")

#load functions 
source("Thesis functions.R")

#load DtD results
load("DtD_results.Rdata")
load("Reg_Data_Wins.Rdata")


#set wd to data folder
setwd("C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/3. Data (2010)/1. Extracted Data")
list.files()

#load stock prices
prc_load<-load_thesis_data("Stock prices (adjusted).xlsx","Price","%m-%d-%Y","Overview.xlsx")
prc<-prc_load$df
  
#load scope data 
scope1_load<-load_thesis_data("Scope 1.xlsx","Scope1","%m-%d-%Y","Overview.xlsx")
scope1<-scope1_load$df

#load scope data 
scope2_load<-load_thesis_data("Scope 2.xlsx","Scope2","%m-%d-%Y","Overview.xlsx")
scope2<-scope2_load$df

#drop uneccessary variables
rm(prc_load,scope1_load,scope2_load)

##############################
#DATA PREP
##############################

#setting WD to code folder  
setwd("C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#get initial firm list 
firm_list<-initial_firms(prc,"Price")

#filter data down to relevant firms and period
scope12<-Data_Wins%>%
  select(Firm,Date,Scope12R)

####################################################
#Create plot of DtD vs. scope 1 by quantiles 
####################################################

#merge DtD and scope 1 (dropping all DtD observations where scope 1 is NA)
DtD_vs_scope1<-merge(DtD.results,scope1,by=c("Firm","Date"))%>%
  #arrange data
  arrange(as.numeric(Firm),Date)%>%
  #create year column 
  mutate(Year=year(Date))%>%
  #group by year
  group_by(Year)%>%
  #divide data into quantiles by year 
  mutate(quantile_bin=ntile(Scope1,5))%>%
  #divide data into quartile by year
  mutate(quartile_bin=ntile(Scope1,4))%>%
  #divide into top+bottom 
  mutate(half_bin=ntile(Scope1,2))%>%
  ungroup()


#determine summary statistics based on the quantiles being defined by year 
stats_DtD_vs_scope1<-DtD_vs_scope1%>%
  group_by(quantile_bin)%>%
  summarise(
    Count=n(),
    Mean= mean(DtD, na.rm = TRUE),
    Median = median(DtD, na.rm = TRUE),
    SD = sd(DtD, na.rm = TRUE),
    Min = min(DtD, na.rm = TRUE),
    Max = max(DtD, na.rm = TRUE),
    IQR = IQR(DtD, na.rm = TRUE)
  )%>%
  ungroup()

########################################
#Create DtD vs. Scope 1+2 by quantiles  
########################################

#merge DtD and scope 1 (dropping all DtD observations where scope 1 is NA)
DtD_vs_scope12<-merge(scope12,DtD.results,by=c("Firm","Date"))%>%
  #arrange data
  arrange(as.numeric(Firm),Date)%>%
  #create year column 
  mutate(Year=year(Date))%>%
  #group by year
  group_by(Year)%>%
  #divide data into quantiles by year 
  mutate(quantile_bin=ntile(Scope12R,5))%>%
  #divide into quartiles by year 
  mutate(quartile_bin=ntile(Scope12R,4))%>%
  #divide into top an bottom 
  mutate(half_bin=ntile(Scope12R,2))%>%
  filter(!is.na(Scope12R))%>%
  ungroup()

#determine summary statistics based on the quantiles being defined by year 
stats_DtD_vs_scope12<-DtD_vs_scope12%>%
  group_by(quantile_bin)%>%
  summarise(
    Count=n(),
    Mean= mean(DtD, na.rm = TRUE),
    Median = median(DtD, na.rm = TRUE),
    SD = sd(DtD, na.rm = TRUE),
    Min = min(DtD, na.rm = TRUE),
    Max = max(DtD, na.rm = TRUE),
    IQR = IQR(DtD, na.rm = TRUE)
  )%>%
  ungroup()

#########################################################
#Plot stats 
#########################################################

par(mfrow=c(1,2))

p1<-ggplot(stats_DtD_vs_scope1, aes(x = factor(quantile_bin), y = Mean)) +
  geom_bar(stat = "identity",fill="blue") +
  labs(x = "Quantile", y = "DtD", title = "Average DtD vs.Scope 1") +
  theme_minimal()+
  coord_cartesian(ylim = c(5, 8))+
  scale_y_continuous(breaks = seq(5, 8, by = 0.5)) 

loadfonts()

p2<-ggplot(stats_DtD_vs_scope12, aes(x = factor(quantile_bin), y = Mean)) +
  geom_bar(stat = "identity",fill="darkblue") +
  labs(x = "Quantile", y = "Distance to Default") +
  theme_minimal()+
  coord_cartesian(ylim = c(5, 7.5))+
  scale_y_continuous(breaks = seq(5, 7.5, by = 0.5))+
  theme(
  text = element_text(family = "CMU Serif"),
  panel.grid.major = element_blank(),  # Removes major grid lines
  panel.grid.minor = element_blank(),  # Removes minor grid lines
  axis.line.y = element_line(colour = "black"), # Ensures axis lines are visible
  axis.line.x = element_line(colour = "black"),
  axis.ticks.x=element_blank())


plot(p2)

ggsave("DtD vs Scope.png", plot = p2, width = 5, height = 3, dpi = 300)


###########################################################################
#create stats for DtD vs. Scope 1 by quantile by year and for Scope 1+2 
###########################################################################

#determine summary statistics based on the quantiles being defined by year 
stats_DtD_vs_scope1_by_year<-DtD_vs_scope1%>%
  group_by(Year,quartile_bin)%>%
  summarise(
    Count=n(),
    Mean= mean(DtD, na.rm = TRUE),
    Median = median(DtD, na.rm = TRUE),
    SD = sd(DtD, na.rm = TRUE),
    Min = min(DtD, na.rm = TRUE),
    Max = max(DtD, na.rm = TRUE),
    IQR = IQR(DtD, na.rm = TRUE)
  )%>%
  ungroup()


#determine summary statistics based on the quantiles being defined by year 
stats_DtD_vs_scope12_by_year<-DtD_vs_scope12%>%
  group_by(Year,quartile_bin)%>%
  summarise(
    Count=n(),
    Mean= mean(DtD, na.rm = TRUE),
    Median = median(DtD, na.rm = TRUE),
    SD = sd(DtD, na.rm = TRUE),
    Min = min(DtD, na.rm = TRUE),
    Max = max(DtD, na.rm = TRUE),
    IQR = IQR(DtD, na.rm = TRUE)
  )%>%
  ungroup()

###################################################
#plot stats
###################################################


#Plotting
p3<-ggplot(stats_DtD_vs_scope1_by_year, aes(x = Year, y = Mean, color = as.factor(quartile_bin))) +
  geom_line() + # Add line plot
  geom_point() + # Optionally add points
  scale_x_continuous(breaks = function(x) pretty(x, n = 10, min.n = 1, only.loose = TRUE, tight.log = FALSE, high.u.bias = 1.5, u5.bias = 0.5)) +
  labs(title = "Average DtD vs. Scope 1",
       x = "Year",
       y = "Average DtD",
       color = "Quartile") +
  theme_minimal() # Using a minimal theme for aesthetics


#Plotting
p4<-ggplot(stats_DtD_vs_scope12_by_year, aes(x = Year, y = Mean, color = as.factor(quartile_bin))) +
  geom_line() + # Add line plot
  geom_point() + # Optionally add points
  scale_x_continuous(breaks = function(x) pretty(x, n = 10, min.n = 1, only.loose = TRUE, tight.log = FALSE, high.u.bias = 1.5, u5.bias = 0.5)) +
  labs(title = "Average DtD vs. Scope 1+2",
       x = "Year",
       y = "Average DtD",
       color = "Quartile") +
  theme_minimal() # Using a minimal theme for aesthetics

grid.arrange(p3,p4,ncol=2)


###################################################
#Top and bottom 
##################################################

#determine summary statistics based on the quantiles being defined by year 
stats_DtD_vs_scope1_by_year_tb<-DtD_vs_scope1%>%
  group_by(Year,half_bin)%>%
  summarise(
    Count=n(),
    Mean= mean(DtD, na.rm = TRUE),
    Median = median(DtD, na.rm = TRUE),
    SD = sd(DtD, na.rm = TRUE),
    Min = min(DtD, na.rm = TRUE),
    Max = max(DtD, na.rm = TRUE),
    IQR = IQR(DtD, na.rm = TRUE)
  )%>%
  ungroup()


#determine summary statistics based on the quantiles being defined by year 
stats_DtD_vs_scope12_by_year_tb<-DtD_vs_scope12%>%
  group_by(Year,half_bin)%>%
  summarise(
    Count=n(),
    Mean= mean(DtD, na.rm = TRUE),
    Median = median(DtD, na.rm = TRUE),
    SD = sd(DtD, na.rm = TRUE),
    Min = min(DtD, na.rm = TRUE),
    Max = max(DtD, na.rm = TRUE),
    IQR = IQR(DtD, na.rm = TRUE)
  )%>%
  ungroup()


#Plotting
p5<-ggplot(stats_DtD_vs_scope1_by_year_tb, aes(x = Year, y = Mean, color = as.factor(half_bin))) +
  geom_line() + # Add line plot
  geom_point() + # Optionally add points
  scale_x_continuous(breaks = function(x) pretty(x, n = 10, min.n = 1, only.loose = TRUE, tight.log = FALSE, high.u.bias = 1.5, u5.bias = 0.5)) +
  labs(title = "Average DtD vs. Scope 1",
       x = "Year",
       y = "Average DtD",
       color = "Top and bottom") +
  theme_minimal() # Using a minimal theme for aesthetics


#Plotting
p6<-ggplot(stats_DtD_vs_scope12_by_year_tb, aes(x = Year, y = Mean, color = as.factor(half_bin))) +
  geom_line() + # Add line plot
  geom_point() + # Optionally add points
  scale_x_continuous(breaks = function(x) pretty(x, n = 10, min.n = 1, only.loose = TRUE, tight.log = FALSE, high.u.bias = 1.5, u5.bias = 0.5)) +
  labs(title = "Average DtD vs. Scope 1+2",
       x = "Year",
       y = "Average DtD",
       color = "Top and bottom") +
  theme_minimal() # Using a minimal theme for aesthetics

grid.arrange(p5,p6,ncol=2)

#save all plots

# Open a PDF device
pdf("DtD vs. Scope.pdf", width = 7, height = 5) # Adjust size as needed

# Create your plots here. For example:
print(p1)
print(p2)
print(p3)
print(p4)
print(p5)
print(p6)

# Close the device
dev.off()


###############################
#DETERMINE END OF YEAR MCAP (e)
###############################

#merge price and shares outstanding by date and firm 
equity_data <- merge(prc, shrout, by = c("Date", "Firm")) %>%
  #make column defining observation year
  mutate(Year=year(ymd(Date)))%>%
  #group data by year and then by firm 
  group_by(Year,Firm)%>%
  #Take the last day within each group i.e., last date of year
  slice_max(order_by=Date,n=1)%>%
  #sort merged data by firm and then date
  arrange(Firm, Date)%>%
  #calculate market cap
  mutate(e = Price * Shrout)%>%
  #keep only relevant variables 
  select(Date,Firm,e)

#drop prc and shrout 
rm(prc,shrout)

###################################
#DETERMINE WEIGHTS (w) 
###################################

#Initially just determine the weights based on 2022 EoY mcap, consider rebalancing each year
#determine total mcap in 2022
mcap_total_2022<-sum(equity_data$e[year(equity_data$Date)==2022])

weights<-equity_data %>%
  #add total mcap in 2022 to dataframe
  mutate(mcap_total_2022=mcap_total_2022)%>%
  #group data by year
  group_by(Year)%>%
  #sum the mcap of all firms by year (needed if we want to rebalance every year)
  mutate(mcap_total=sum(e))%>%
  #ungroup 
  ungroup()%>%
  #determine weights based on 2022
  mutate(w_2022=e/mcap_total_2022)%>%
  #determine weights with annual rebalancing
  mutate(w=e/mcap_total)%>%
  #select only needed variables
  select(Date,Firm,w_2022,w,Year)

######################################
#create final data set
######################################

DtD_vs_scope<-merge((merge(DtD.results,weights,by=c("Firm","Date"))),scope1,by=c("Firm","Date"))%>%
  #determine weighted DtD (based on 2022 weigths)
  mutate(w_2022_DtD=w_2022*DtD)%>%
  #determine weigthed DtD with rebalancing 
  mutate(w_DtD=w*DtD)%>%
  #determine weighted Scope 1 (based on 2022 weigths)
  mutate(w_2022_scope1=w_2022*Scope1)%>%
  #determine weighted Scope 1 with rebalancing 
  mutate(w_scope1=w*Scope1)%>%
  #keep only necessary variables 
  select(Date,Firm,Scope1,DtD,w_2022_DtD,w_DtD,w_2022_scope1,w_scope1,Year,w_2022,w)


##################################################
#Making plot of weighted DtD vs. weighted Scope 1
##################################################

yearly_summary <- aggregate(cbind(w_2022_DtD, w_DtD, w_2022_scope1, w_scope1,w,w_2022) ~ Year, data=DtD_vs_scope, FUN=sum, na.rm=TRUE)

par(mfrow=(c(1,2)))

plot(yearly_summary$Year, yearly_summary$w_DtD, type="l", xlab="Year", ylab="Weighted DtD",col="blue")

plot(yearly_summary$Year, yearly_summary$w_scope1, type="l", xlab="Year", ylab="Weighted Scope1",col="red")

###########################################################
#Making plot of absolute DtD vs. Absolute Scope 1 only 2022 
###########################################################

par(mfrow=(c(1,1)))

DtD_vs_scope_2022<-DtD_vs_scope%>%
  filter(Date==as.Date("2022-12-30"))

plot(DtD_vs_scope_2022$DtD,DtD_vs_scope_2022$Scope1,type="p",xlab="DtD", ylab="Scope 1",ylim=c(20831,452881))

summary(DtD_vs_scope_2022)

##########################################################
#Making plot of DtD vs. Scope 1 intensity 
##########################################################

#temporary code to create mergedate to deal with wrong date in data 
revenue<-revenue%>%
  mutate(mergedate=format(as.Date(Date),"%Y-%m"))%>%
  select(mergedate,Revenue,Firm)
  
DtD_vs_scope_intensity.v1<-merge(DtD.results,scope1,by=c("Date","Firm"))%>%
  mutate(mergedate=format(as.Date(Date),"%Y-%m"))

DtD_vs_scope_intensity<-merge(DtD_vs_scope_intensity.v1,revenue,by=c("mergedate","Firm"))%>%
  arrange(Firm,Date)%>%
  mutate(scope1_intensity=Scope1/(Revenue))%>%
  select(Date,Firm,DtD,Scope1,scope1_intensity,Revenue)

attach(DtD_vs_scope_intensity)
plot(Revenue,Scope1,type="p",xlab="Revenue",ylab="Scope 1")


summary(DtD_vs_scope_intensity)

plot(scope1_intensity,DtD,type="p",xlab="Scope 1 intensity",ylab="DtD")

plot(Firm,scope1_intensity,type="p",xlab="Firm",ylab="Scope 1 intensity")

?left_join

