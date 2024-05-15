
######################################
#Load, rename, and sort files function
######################################

##EXAMPLE
##load shares outstanding 
##shrout_load<-load_thesis_data("Shares outstanding.xlsx","Shrout","%m-%d-%Y","Overview.xlsx")


load_thesis_data<-function(file_path,value_name,date_format,overview_path){
  #load main excel and define NA as NA values
  df<-read_excel(file_path,na=c("NA",0))
  
  #load the fake firm ID's
  Overview<-read_excel(overview_path)
  
  #name first column as Date
  names(df)[1]<-"Date"
  df$Date<-as.Date(df$Date,format=date_format)
  
  #create vector with all firms in the main excel file and match with firms in overview
  col_names<-names(df)[-1]
  matched_firms<-match(col_names,Overview$Name)
  
  #create list of unmatched firms
  unmatched_firms<-col_names[is.na(matched_firms)]
  
  #prepare for changing from firm names to fake firm ID's 
  new_names<-ifelse(is.na(matched_firms),col_names,Overview$Firm[matched_firms])
  
  #remove columns without a matching fake firm ID in overview
  df<-df%>%
    select(Date,col_names[!is.na(matched_firms)])
  
  #rename columns 
  names(df)[-1]<-new_names[!is.na(matched_firms)]
  
  #convert column names to character to make the reshaping work
  df[,-1] <- lapply(df[,-1], as.character)
  
  #reshape data and arrange by Date and Firm
  df<-df%>%
    pivot_longer(
      cols = -Date,
      names_to = "Firm",
      values_to = value_name,
      ) %>%
    arrange(as.numeric(Firm),Date)
  
  df$Firm<-as.numeric(df$Firm)
  
  #define the loaded variable as numeric and changes 0 values to NA 
  df[[value_name]] <- as.numeric(df[[value_name]])
  df[[value_name]][df[[value_name]] == 0] <- NA
  
  return(list(df=df,unmatched_firms=unmatched_firms))
  
}

########################################
#initial firm list 
########################################

initial_firms<-function(stockdata,price_col,threshold=252){
  #threshold default is 1 years of stockdata change if needed
  data<-stockdata%>%
    group_by(Firm)%>%
    #create column that accumulates the number of non missing stockprices
    mutate(consecutive_non_na=cumsum(!is.na(.data[[price_col]])))%>%
    #removes data with less than threshold number of observations
    filter(max(consecutive_non_na)>=threshold)%>%
    ungroup()%>%
    select(-consecutive_non_na)
  
  firm_list<-distinct(data,Firm)
  }

###########################################
#sort data initial period and initial firms
###########################################

filter_firms_and_date <- function(data, start_year, end_dates_data, selected_firms) {
  # Ensure start_year is of type integer
  start_year <- as.integer(start_year)
  
  # Prepare the end dates data
  # Assuming the end_dates_data has columns named "Firm" and "EndDate" with EndDate being a year
  end_dates_data <- end_dates_data %>%
    mutate(EndDate = as.integer(EndDate))
  
  # Merge selected firms with end dates data to include the custom end year
  selected_firms_with_end_dates <- selected_firms %>%
    left_join(end_dates_data, by = "Firm")
  
  # Filter data based on selected firms and dynamic end date range
  filtered_data <- data %>%
    full_join(selected_firms_with_end_dates, by = "Firm") %>%
    mutate(Year = year(Date)) %>%
    rowwise() %>%
    filter(Year >= start_year, Year <= EndDate) %>%
    ungroup() %>%
    select(-EndDate)%>%
    arrange(as.numeric(Firm), Date)
  
  # Remove missing items
  complete_data <- na.omit(filtered_data)
  
  return(complete_data)
}

#############################################
#get riskfree rate 
#############################################

get_rf<-function(file_path,start_year,end_year,date_format){
  df<-read_excel(file_path)
  names(df)[1]<-"Date"
  df$Date<-as.Date(df$Date,format=date_format)
  names(df)[2]<-"r"
  
  start_year <- as.integer(start_year)
  end_year <- as.integer(end_year)
  
  #Filter data based on selected firms and date range
  filtered_data <- df %>%
    filter(year(Date) >= start_year, year(Date) <= end_year)%>%
    arrange(Date)
  
  #Define the rate as numerical data and replace 0 values with NA 
  filtered_data[[2]] <- as.numeric(filtered_data[[2]])/100
  
  return(filtered_data)
}

#############################################
#get macro data
#############################################

get_macro<-function(file_path,value_name,start_year,end_year,date_format){
  df<-read_excel(file_path)
  names(df)[1]<-"Date"
  df$Date<-as.Date(df$Date,format=date_format)
  names(df)[2]<-value_name
  
  start_year <- as.integer(start_year)
  end_year <- as.integer(end_year)
  
  #Filter data based on selected firms and date range
  filtered_data <- df %>%
    filter(year(Date) >= start_year, year(Date) <= end_year)%>%
    arrange(Date)
  
  filtered_data[[value_name]]<-as.numeric(filtered_data[[value_name]])
  filtered_data[[value_name]][filtered_data[[value_name]]==0]<-NA
  
  return(filtered_data)
}

##############################################
#functions to compute Mertons DtD
##############################################

#create the objective function that is to be minimized using the BSM formulas 
objective_func<-function(a,e,f,r,va,T){
  d1<-(log(a/f)+(r+va^2/2)*T)/(va*sqrt(T))
  d2<-(log(a/f)+(r-va^2/2)*T)/(va*sqrt(T))
  e.calculated<-a*pnorm(d1)-f*exp(-r*T)*pnorm(d2)
  return((e-e.calculated)^2)
}

#create function to perform the optimization solving for a 
solve_for_a<-function(e,f,r,va,T,guess){
  #perform the optimization to solve 
  a.result<-optim(
    par=guess,
    fn=objective_func,
    e=e,
    f=f,
    r=r,
    va=va,
    T=T,
    method="BFGS"
  )
  return(a.result$par)
}


