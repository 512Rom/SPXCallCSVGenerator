library(openxlsx);library(Quandl);library(quantmod);library(lubridate);library(Quandl);library(quantmod)

DailyCalls<-function(){
  TodaysOptions<-createWorkbook() # creates WB
  Alloptions<-getOptionChain('^SPX','2022/2023') # Loads data
  addWorksheet(TodaysOptions, sheetName = paste("Calls",today())) # adds Worksheet "Calls" and todays date
  Calls<-c()
  for(i in 1:length(Alloptions)){
    Calls<-rbind(Calls,Alloptions[[i]]$calls)
    
  }
  writeData(TodaysOptions, paste("Calls",today()), x = Calls)
  saveWorkbook(TodaysOptions, paste(today(),"Calls.xlsx"))
}


DailyPuts<-function(){
  TodaysOptions<-createWorkbook()
  Alloptions<-getOptionChain('^SPX','2022/2023')
  addWorksheet(TodaysOptions, sheetName = paste("Puts",today()))
  Puts<-c()
  for(i in 1:length(Alloptions)){
    Puts<-rbind(Puts,Alloptions[[i]]$puts)
    
  }
  writeData(TodaysOptions, paste("Puts",today()), x = Puts)
  saveWorkbook(TodaysOptions, paste(today(),"Puts.xlsx"))
}

DailyCalls()
DailyPuts()
