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

#setting WD to code folder  
setwd("C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#Loading DtD
load("DtD_wip.Rdata")
load("DtDdata_wip.Rdata")
load("Reg_Data_Wins2.Rdata")


####################################################
# Remove obs of DtD for firms with insufficient data
####################################################

final_obs<-Data_Wins2%>%
  select(Date,Firm)

DtD_wip_filtered<-semi_join(DtD_wip,final_obs,by=c("Date","Firm"))%>%
  filter(Firm!=226)

###############################################
#Examine DtD
###############################################

#start by plotting data 

# Add an index column to the dataframe 
DtD_wip_filtered$Index <- seq_along(DtD_wip_filtered$DtD)

# make scatterplot
outlier_plot<-ggplot(DtD_wip_filtered, aes(x = Index, y = DtD)) +
  geom_jitter(width = 0.2, height = 0, color = "black") +  # Adjust jitter width as needed
  geom_point(data = subset(DtD_wip_filtered, DtD > 25),
             aes(x = Index, y = DtD), color = "red",size=2.5) +  # Mark outliers in red
  labs(       y = "Distance to Default",
              x = "Observation number") +
  geom_hline(yintercept = 0, linetype = 1) +
  theme_minimal() +
  theme(
    text = element_text(family = "CMU Serif"),
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.line.y = element_line(colour = "black"), # Ensures axis lines are visible
    axis.line.x=element_blank(),
    axis.ticks.x=element_blank(),
    axis.text.x=element_blank())+
  scale_x_continuous(breaks = seq(0, max(DtD_wip_filtered$Index), by = 1000), limits = c(0, max(DtD_wip_filtered$Index)*1.01),expand=c(0,0.5)) +
  scale_y_continuous(breaks = seq(-5, 55, by = 5), limits = c(-5, 55))

#make summary statistics 
summary_stats_DtD <- DtD_wip_filtered %>%
  summarise(
    n = n(),
    mean_DtD = mean(DtD, na.rm = TRUE),
    median_DtD = median(DtD, na.rm = TRUE),
    min_DtD = min(DtD, na.rm = TRUE),
    max_DtD = max(DtD, na.rm = TRUE),
    sd_DtD = sd(DtD, na.rm = TRUE)
  )


ggsave("DtD outlier plot.png", plot = outlier_plot, width = 6, height = 4, dpi = 300)

plot(outlier_plot)



print(summary_stats_DtD)

summary_stats_DtD_table <- xtable(summary_stats_DtD)

# Print the LaTeX code
print(summary_stats_DtD_table, type = "latex", include.rownames = FALSE)


#########################################
#Examine outliers
#########################################

outlier_data<-DtD.data%>%
  filter(Firm==269 | Firm==323)%>%
  mutate(Date=as.Date(Date,format="%Y/%m/%d"))

outlier_table1<-outlier_data%>%
  filter(DtD>=25)%>%
  select(Firm,Date,DtD,a.hat,f,mu,assetvol)

outlier_table2<-outlier_data%>%
  na.omit%>%
  group_by(Firm)%>%
  summarise(
    mean_DtD = mean(DtD, na.rm = TRUE),
    median_DtD = median(DtD, na.rm = TRUE),
    min_DtD = min(DtD, na.rm = TRUE),
    max_DtD = max(DtD, na.rm = TRUE),
    sd_DtD = sd(DtD, na.rm = TRUE),
    n = n()  
  )

outlier_fig_data<-outlier_data%>%
  na.omit%>%
  mutate(Date=year(Date))
 
#plot: 
par(mfrow=c(1,1))

# Create the line plot
outlier_timeseries_plot<-ggplot(data = outlier_fig_data, aes(x = Date, y = DtD, group = Firm, color = as.factor(Firm))) +
  geom_line(size=0.6) +  # Adds lines to connect points for each firm
  geom_point() + # Adds points to highlight each data point
  labs(
       x = "Year",
       y = "Distance to Default",
       color = "Firm") +
  theme_minimal()+
  theme(
    text = element_text(family = "CMU Serif"),
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.line = element_line(colour = "black")  
  ) +
  scale_color_manual(values = c("darkblue", "lightblue"))+  
  scale_x_continuous(breaks = seq(2010, 2022, by = 1), limits = c(2010,2022)) +
  scale_y_continuous(breaks = seq(0, 55, by = 5), limits = c(0, 57),expand=c(0,0))


ggsave("DtD outlier timeseries plot.png", plot = outlier_timeseries_plot, width = 6, height = 4, dpi = 300)

##############################
#Negative firms 
##############################

#make list of firms with negative DtD
negative_firms<-DtD.data%>%
  filter(DtD<=0)%>%
  filter(Firm!=226)

#Extract all data related to firms with negative DtD
negative_data<-DtD.data%>%
  filter(Firm %in% negative_firms$Firm)%>%
  mutate(Date=as.Date(Date,format="%Y/%m/%d"))%>%
  na.omit%>%
  filter(Firm!=226)

negative_fig_data <- negative_data %>%
  mutate(Date = as.numeric(format(Date, "%Y")),  # Extract year from Date
         Firm = factor(Firm, levels = c(setdiff(unique(Firm), c(66, 187, 421, 134, 222, 260)),
                                        134, 222, 260, 
                                        66, 187, 421)))

#top row firms: 66,187,421,226

negative_firms_timeseries_plot<-ggplot(data = negative_fig_data, aes(x = Date, y = DtD, color = as.factor(Firm))) +
  geom_line(colour="black") +  # Adds lines to connect points for each firm
  geom_point(data = subset(negative_fig_data, DtD < 0),
             aes(x = Date, y = DtD), size=1.5, colour="red") +
  labs(
    x = "Year",
    y = "Distance to Default",
    color = "Firm") +
  geom_hline(yintercept = 0, linetype = 2) +
  theme_minimal() +
  theme(
    text = element_text(family = "CMU Serif"),
    panel.grid.major = element_blank(),  # Removes major grid lines
    panel.grid.minor = element_blank(),  # Removes minor grid lines
    axis.line = element_line(colour = "black")  
  ) +
  scale_x_continuous(breaks = c(2010, 2022), limits = c(2010, 2022)) +
  scale_y_continuous(breaks = 0, limits = c(-3, 8)) +
  facet_wrap(~ Firm, scales = "free_y",  nrow = 5, ncol = 3, labeller = labeller(Firm = function(x) paste("Firm", x)))
  
ggsave("Negative firms.png", plot = negative_firms_timeseries_plot, width = 8, height = 6, dpi = 300)

#Create table for Appendix
negative_firms_table_data <- negative_data %>%
  filter(DtD <= 0) %>%
  select(Firm,Date, DtD, a.hat, f, mu, assetvol) %>%
  mutate(
    Date = as.integer(year(Date)), 
    Firm = as.factor(Firm),  
    #a.hat = formatC(a.hat, format = "f", big.mark = ",", digits = 0), 
    #f = formatC(f, format = "f", big.mark = ",", digits = 0), 
    "V_t/f"=a.hat/f
  )%>%
  select(Firm,Date,DtD,"V_t/f",mu,assetvol)

negative_firms_table<- xtable(negative_firms_table_data)

# Print the LaTeX code
print(negative_firms_table, type = "latex", include.rownames = FALSE,
      digits = c(0, 0, 2, 0, 0, 2, 2))

###############################################################################
#Make final data set
###############################################################################

Clean_DtD_neg_zero<-DtD_wip%>%
  mutate(DtD=ifelse(DtD<=0,0,DtD))%>%
  filter(DtD<=50)

Clean_DtD_neg_rm<-DtD_wip%>%
  filter(DtD>=0&DtD<=50)

print(summary(Clean_DtD_neg_zero$DtD))
print(summary(Clean_DtD_neg_rm$DtD))

save(Clean_DtD_neg_zero,file="Clean_DtD_neg_zero.Rdata")
save(Clean_DtD_neg_rm,file="Clean_DtD_neg_rm.Rdata")




check<-DtD_wip_filtered%>%
  filter(Firm==269)

check1<-DtD.data%>%
  filter(Firm==269)%>%
  na.omit

