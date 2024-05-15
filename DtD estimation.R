#loading packages
library(readxl)
library(xts)
library(tidyr)
library(dplyr)
library(zoo)
library(lubridate)
library(DtD)
library(ggplot2)

#removing all previous variables 
rm(list=ls())

#setting WD to code folder  
setwd("C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#load functions 
source("Thesis functions.R")

#######################################
#load data
#######################################

#setting WD to data folder
setwd("C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/3. Data (2010)/1. Extracted Data")
list.files()

#loading data files, reshaping them and sorting by date and firm 
prcload<-load_thesis_data("Stock prices (adjusted).xlsx","Price","%m/%d/%Y","Overview.xlsx")
prc<-prcload$df

prc_mktcapload<-load_thesis_data("Stock prices (mktcap).xlsx","Price","%m/%d/%Y","Overview.xlsx")
prc_mktcap<-prc_mktcapload$df

shroutload<-load_thesis_data("Shares outstanding.xlsx","Shrout","%m/%d/%Y","Overview.xlsx")
shrout<-shroutload$df

stdebtload<-load_thesis_data("Short term liabilities quarterly EUR.xlsx","STdebt","%m/%d/%Y","Overview.xlsx")
stdebt<-stdebtload$df

totaldebtload<-load_thesis_data("Total liabilities quarterly EUR.xlsx","Totaldebt","%m/%d/%Y","Overview.xlsx")
totaldebt<-totaldebtload$df

listings<-read_excel("Listing.xlsx")

rf<-get_rf("Euribor 12 m.xlsx",2010,2022,"%m/%d/%Y")

#drop load dataframes
rm(list = ls()[grepl("load", ls())])

#setting WD to code folder  
setwd("C:/Users/Veron/OneDrive - CBS - Copenhagen Business School/Thesis/4. R code")
list.files()

#########################################
#Get initial firm list 
#########################################

firm_list<-initial_firms(prc,"Price")

#########################################
#sort data
#########################################

#remove data prior to 2010, after potential delistings, firms not included in the firm list and NA values
prc<-filter_firms_and_date(prc,2010,listings,firm_list)

prc_mktcap<-filter_firms_and_date(prc_mktcap,2010,listings,firm_list)

shrout<-filter_firms_and_date(shrout,2010,listings,firm_list)

stdebt<-filter_firms_and_date(stdebt,2010,listings,firm_list)

totaldebt<-filter_firms_and_date(totaldebt,2010,listings,firm_list)

###########################
#Determining Market Cap (e)
###########################

#merge price and shares outstanding by date and firm 
merged.equity.data <- merge(prc_mktcap, shrout, by = c("Date", "Firm")) %>%
  # sort merged data by firm and then date 
  arrange(Firm, Date) %>%
  # calculate market cap
  mutate(e = Price * Shrout)%>%
  # create short date 
  mutate(mergedate=paste(year(Date),quarter(Date),sep="-Q"))%>%
  # remove the unadjusted stock price 
  select(-Price)

#add the adjusted share prices to the dataframe
merged.equity.data<-merge(merged.equity.data,prc,by=c("Date","Firm"))

#drop prc and shrout 
rm(prc,prc_mktcap,shrout)

####################################################
#Determing long term liabilities and face value (f)
####################################################

#merge ST and LT debt in one data frame
merged.debt.data<-merge(totaldebt,stdebt,by=c("Date","Firm"))%>%
  #sort by firm then date 
  arrange(Firm,Date)%>%
  #calculate long term debt as total debt - short term debt
  mutate(LTdebt=Totaldebt-STdebt)%>%
  #set any negative debt values equal to zero
  mutate(LTdebt=pmax(LTdebt,0))%>%
  mutate(STdebt=pmax(STdebt,0))%>%
  #determine f as STdebt+0.5*LTdebt 
  mutate(f=STdebt+0.5*LTdebt)%>%
  #create merge date 
  mutate(mergedate=paste(year(Date),quarter(Date),sep="-Q"))%>%
  #remove the original "Date" column
  select(-Date)

#drop total debt and short term debt 
rm(stdebt, totaldebt)

#####################################################
#make combined data set to determine DtD 
#####################################################

#merge the two data sets 
DtD.data<-merge(merge(merged.equity.data,merged.debt.data,by=c("mergedate","Firm"),all.x=TRUE),rf,by="Date")%>%
  #arrange by Firm and by the daily date from the daily stock returns 
  arrange(Firm,Date) %>%
  # determine a as Mcap+FV
  mutate(a=e+f)%>%
  #Keep only necessary variables
  select(Date,Firm, e, f, a,r,Price)%>%
  #drop missing variables 
  na.omit()

sum(is.na(DtD.data))

#drop merge data sets 
rm(list = ls()[grepl("merge|rf", ls())])


###################################
#Calculate initial asset volatility 
###################################

DtD.data<-DtD.data%>%
  #arranging by Date and Firm 
  arrange("Date","Firm")%>%
  #group by Firm so following calculations are done by firm 
  group_by(Firm)%>%
  #determine daily log return for each firm 
  mutate(re=log(Price/lag(Price)))%>%
  #determine the equity volatility based on a 252-days rolling window (UPDATE if more than 1 year)
  mutate(ve=rollapply(re,width=252,FUN=sd,fill=NA,align="right"))%>%
  #determine the annualized equity volatility 
  mutate(ve.annual=ve*sqrt(252))%>%
  #determine the initial asset volatility value using the formula va=ve.annual*e/a
  mutate(va=ve.annual*e/a)%>%
  #ungroup since firm specific computations are now done 
  ungroup()%>%
  #create forecasting time horizon variable 
  mutate(T=1)%>%
  #add empty column to store values of DtD
  mutate(DtD=NA,mu=NA,assetvol=NA,a.hat=NA)%>%
  #sort data by firm and then by date 
  arrange(Firm,Date)%>%
  #drop unnecessary variables 
  select(-c(re,ve.annual,ve))

##############################################################
#Iteration step 1: make data set with last observation of year
##############################################################

#defining the dates as dates in R 
DtD.data$Date<-as.Date(DtD.data$Date)

#make a dataframe containing the date of the last observation of the year and the corresponding initial asset volatility for each firm in the data 
last.day<-DtD.data%>%
  #making a column defining the observation year
  mutate(Year = year(ymd(Date))) %>%
  #grouping data by year and then by firm
  group_by(Year,Firm)%>%
  #takes the max date within each grouping i.e., last trading day of the year
  summarize(Date=max(Date))%>%
  #drop dates that are not last trading day of the year for companies delisted througout the year 
  filter(month(Date)==12)%>%
  #ad the corresponding initial asset volatility to the last day data
  left_join(DtD.data%>%select(Firm,Date,va),by=c("Firm","Date"))%>%
  #arrange by Firm and then Date
  arrange(Firm,Date)%>%
  #drop observations with missing initial asset volatility values 
  na.omit(last.day$va)

#removing the va column from the DTD.data
DtD.data<-DtD.data%>%
  select(-va)

###########################################################
#Iteration step 2: extract list of firms in the data sample
###########################################################

#extract the number of firms in the sample 
unique.firms<-unique(last.day$Firm)

###########################################################
#Iteration step 3: perform iteration 
###########################################################

#tells r that the iteration has to be done for each firm in the sample
for (firm in unique.firms) {
  #make a new data set with the initial asset volatilities and corresponding dates for the firm being looped over 
  loop.firm.data<-last.day[last.day$Firm==firm,c("Date","va")]
  
  #make a loop that will be repeated for all the available last dates for the specific firm 
  for (i in 1:nrow(loop.firm.data)){
    #extract the observation date that is being looped over 
    loop.observation.date<-loop.firm.data$Date[i]
    #extract the corresponding initial asset volatility 
    loop.va<-loop.firm.data$va[i]
    
    
    #performing the actual iteration to estimate va and mu 
    #create convergence flag that ensures iteration continues until convergence of the asset volatility is TRUE
    loop.converged<-FALSE
    while(!loop.converged){
      
      #extract the preceeding 252 trading days of observations for the firm on the date being looped over 
      loop.past.data<-DtD.data%>%
        #filter by the firm and takes only observations prior to the observation.date
        filter(Firm==firm&Date<loop.observation.date)%>%
        #now filters down to the 252 prior trading days 
        slice_max(order_by=Date,n=252)%>%
        #add the initial volatility to the past.data 
        mutate(va=loop.va)%>%
        #add a column to store estimated asset values in: 
        mutate(estimated.a=NA,logreturn.a=NA,l.estimated.a=NA)%>%
        #change the dates so earliest date first and latest date last 
        arrange(desc(Date))
      
      #make a loop that estimates the value of a and corresponding asset log returns for the 252 preceding trading days 
      for (row in 1:nrow(loop.past.data)) {
        #extract the corresponding variables from loop.past.data needed to compute asset value
        loop.f <- loop.past.data$f[row]
        loop.r <- loop.past.data$r[row]
        loop.e<-loop.past.data$e[row]
        loop.T<-loop.past.data$T[row]
        loop.a<-loop.past.data$a[row]
        loop.va<-loop.past.data$va[row]
        #estimate the asset value for the specific day using the relevant data 
        loop.past.data$estimated.a[row]<-solve_for_a(loop.e,loop.f,loop.r,loop.va,loop.T,guess=loop.a)
        
        #compute the asset log return
        #first setting the logreturn equal to NA and assigning the current asset value as the lagged asset value next loop
        if(row==1){
          loop.past.data$logreturn.a<-NA
          loop.past.data$l.estimated.a[row+1]<-loop.past.data$estimated.a[row]
          #computing the return for row=max 
        } else if(row==nrow(loop.past.data)){
          loop.past.data$logreturn.a[row]<-(log(loop.past.data$estimated.a[row])-log(loop.past.data$l.estimated.a[row]))
          #computing the return for all rows>1 and rows<max assigning the current asset value as the lagged value for the next row 
        } else {
          loop.past.data$logreturn.a[row]<-(log(loop.past.data$estimated.a[row])-log(loop.past.data$l.estimated.a[row]))
          loop.past.data$l.estimated.a[row+1]<-loop.past.data$estimated.a[row]
        }}
      
      #estimate mu tilde, the number of returns in the sample corresponds to #obs-1 since no log return for #obs 1
      loop.mu.tilde<-(log(loop.past.data$estimated.a[nrow(loop.past.data)])-log(loop.past.data$estimated.a[1]))/(nrow(loop.past.data)-1)
      
      #estimate new asset volatility 
      loop.va.new<-sd(na.omit(loop.past.data$logreturn.a))*sqrt(252)
      
      #check if the estimated asset volatility has converged or not 
      if (abs(loop.va.new-loop.va)<1e-12){
        #if the difference between the asset volatility is less than 10^-12 change the value of converge to true 
        loop.converged<-TRUE
        loop.va<-loop.va.new
      } else {
        #if the difference isn't less than 10^-12 update the value of va and continue iteration 
        loop.va<-loop.va.new
      }}
    
    #Compute the drift using the formula for the mean when logreturns are assumed to be normal distributed 
    loop.mu.hat<-(loop.mu.tilde*252)+loop.va^2/2
    
    #Compute the asset value on the observation date using the estimated volatility 
    #extract the necessary values from the DtD dataframe 
    loop.f <- DtD.data$f[DtD.data$Firm==firm &DtD.data$Date==loop.observation.date]
    loop.r <- DtD.data$r[DtD.data$Firm==firm &DtD.data$Date==loop.observation.date]
    loop.e<-DtD.data$e[DtD.data$Firm==firm &DtD.data$Date==loop.observation.date]
    loop.T<-DtD.data$T[DtD.data$Firm==firm &DtD.data$Date==loop.observation.date]
    loop.a<-DtD.data$a[DtD.data$Firm==firm &DtD.data$Date==loop.observation.date]
    
    loop.a.hat<-solve_for_a(loop.e,loop.f,loop.r,loop.va,loop.T,guess=loop.a)
    
    #Store the asset volatility and the drift in the DtD dataframe
    DtD.data$assetvol[DtD.data$Firm==firm & DtD.data$Date==loop.observation.date]<-loop.va
    DtD.data$mu[DtD.data$Firm==firm & DtD.data$Date==loop.observation.date]<-loop.mu.hat
    DtD.data$a.hat[DtD.data$Firm==firm & DtD.data$Date==loop.observation.date]<-loop.a.hat
    
    
    #Compute the DtD and store it in the distance to default dataframe
    DtD.data$DtD[DtD.data$Firm==firm & DtD.data$Date==loop.observation.date]<-(log(loop.a.hat)-log(loop.f)+(loop.mu.hat-loop.va^2/2)*loop.T)/(loop.va*sqrt(loop.T))
  }}

#remove all variables generated in loop 
rm(list = ls()[grepl("loop|firm|i|row", ls())])

#create dataframe with results only 
DtD_wip <- DtD.data%>%
  filter(!is.na(DtD))%>%
  select(c(Date,Firm,mu,assetvol,DtD))

save(DtD_wip,file="DtD_wip.Rdata")
save(DtD.data,file="DtDdata_wip.Rdata")
