################################################
# IST687, Final Project
#
# Group Members: Warren Fernandes, Mark Gopani, Jeffrey Stephens
# 
# Date due: 05/13/2021
#
################################################

# Importing the required libraries
library(tidyverse)
library(imputeTS)
setwd("C:/Users/warre/Downloads")
#Loading the data
H1 <- readxl::read_xlsx("H1-Resort.xlsx")
H2 <- readxl::read_xlsx("H2-City.xlsx")
#Viewing the data
View(H1)
View(H2)
#
str(H1)  #28 columns and 40060 rows
str(H2) #28 columns and 79330 rows

library(kernlab)
library(caret)
library(e1071)


# Cleaning and Transforming
# Casting the following char type variables into factors using as.factor() function and fixing data type of Children in H2.
H1$IsCanceled <- as.factor(H1$IsCanceled)
H1$ReservedRoomType <- as.factor(H1$ReservedRoomType)
H1$AssignedRoomType <- as.factor(H1$AssignedRoomType)
H1$DepositType <- as.factor(H1$DepositType)
H1$CustomerType <- as.factor(H1$CustomerType)
H1$Meal <- as.factor(H1$Meal)
H1$IsRepeatedGuest <- as.factor(H1$IsRepeatedGuest)

H2$IsCanceled <- as.factor(H2$IsCanceled)
H2$ReservedRoomType <- as.factor(H2$ReservedRoomType)
H2$AssignedRoomType <- as.factor(H2$AssignedRoomType)
H2$DepositType <- as.factor(H2$DepositType)
H2$CustomerType <- as.factor(H2$CustomerType)
H2$Meal <- as.factor(H2$Meal)
H2$Children <- as.numeric(H2$Children)
H2$IsRepeatedGuest <- as.factor(H2$IsRepeatedGuest)

# Summary of numeric values in the data sets
summary(H1)
summary(H2)

# Replacing the NA values using interpolation function from ImputeTS
table(is.na(H1))
table(is.na(H2))
sum(is.na(H2$Children)) #We can see that there are 4 NA values in the children column of H2 dataset.
H2$Children <- na_interpolation(H2$Children) #Filling the NA values based on the values around them using na_interpolation() function.
sum(is.na(H2$Children)) #We can see that there are no NA values in the children column anymore

# Changing the unknown value into check-out.
table(H1$ReservationStatus) #We can see that there is an unknown/error value(`) in the ReservationStatus column of the H1 dataset.
H1$ReservationStatus[13] <- 'Check-Out' #We are assuming the data to be 'check-out' based on the values around which have a similar ReservationStatusDate.

# Factorizing arrival dates into seasons
H1$month_name <- format(H1$`Arrival Date`, "%b")
H1$month <- as.numeric(format(H1$`Arrival Date`, "%m"))
H1$season <- 'winter'
H1$season[H1$month >= 3 & H1$month <=5] <- 'spring'
H1$season[H1$month >= 6 & H1$month <=8] <- 'summer'
H1$season[H1$month >= 9 & H1$month <=11] <- 'fall'
H1$season <- factor(H1$season, levels = c("summer", "fall", "winter", "spring"))
summary(H1$season)

H2$month_name <- format(H2$`Arrival Date`, "%b")
H2$month <- as.numeric(format(Hotels$`Arrival Date`, "%m"))
H2$season <- 'winter'
H2$season[H2$month >= 3 & H2$month <=5] <- 'spring'
H2$season[H2$month >= 6 & H2$month <=8] <- 'summer'
H2$season[H2$month >= 9 & H2$month <=11] <- 'fall'
H2$season <- factor(H2$season, levels = c("summer", "fall", "winter", "spring"))
summary(H2$season)

# Creating the Visitor Type variable for H1 and H2
H1$VisitorType <- 'Single'
H1$totalVisitors <- H1$Adults + H1$Children + H1$Babies
H1$VisitorType[H1$totalVisitors == 1] <- 'Single'
H2$VisitorType[H2$totalVisitors == 2] <- 'Couple'
H2$VisitorType[H2$totalVisitors > 2] <- 'Family'
H1$VisitorType <- as.factor(H1$VisitorType)
H1$totalVisitors <- NULL
summary(H1$VisitorType)


H2$VisitorType <- 'Single'
H2$totalVisitors <- H2$Adults + H2$Children + H2$Babies
H2$VisitorType[H2$totalVisitors == 1] <- 'Single'
H2$VisitorType[H2$totalVisitors == 2] <- 'Couple'
H2$VisitorType[H2$totalVisitors > 2] <- 'Family'
H2$VisitorType <- as.factor(H2$VisitorType)
H2$totalVisitors <- NULL
summary(H2$VisitorType)

# Calculating the Average Revenue of Stay into a new column
H1$ARS <- (H1$StaysInWeekendNights + H1$StaysInWeekNights) * H1$ADR
H2$ARS <- (H2$StaysInWeekendNights + H2$StaysInWeekNights) * H2$ADR

# Combining the two datasets for comparitive visualizations
H1$HotelType = ("Resort")
H2$HotelType = ("City")
#Hotels = na.omit(Hotels)

Hotels = rbind(H2, H1)
# Average Revenue per stay
#Hotels$ARS <- (Hotels$StaysInWeekendNights + Hotels$StaysInWeekNights) * Hotels$ADR
View(Hotels)

Hotels$ARS
cuts = quantile(Hotels$ARS, c(0, 0.3, 0.6, 0.8, 1))
table(Hotels$RevenueCategory)
Hotels$RevenueCategory <- "Very High"
Hotels$RevenueCategory[Hotels$ARS >= cuts[1] & Hotels$ARS <  cuts[2]]  = "Low"
Hotels$RevenueCategory[Hotels$ARS >= cuts[2] & Hotels$ARS <  cuts[3]]  = "Moderate"
Hotels$RevenueCategory[Hotels$ARS >= cuts[3] & Hotels$ARS <  cuts[4]] = "High"
#Hotels$RevenueCategory[Hotels$ARS >= cuts[4] & Hotels$ARS <=  cuts[5]] = ""
View(Hotels$RevenueCategory)

ggplot(data = Hotels,aes(IsCanceled))+ geom_histogram(stat= "count", binwidth = 0.5, col='black', fill='blue', alpha = 0.4) + facet_wrap(~HotelType)
ggplot(data = Hotels,aes(x= ARS, y= CustomerType))+ geom_bar(stat= 'identity', fill='blue', alpha = 0.4) + facet_wrap(~HotelType)
ggplot(data = Hotels,aes(IsRepeatedGuest))+ geom_histogram(stat= 'count', binwidth = 0.5, col='black', fill='blue', alpha = 0.4) + facet_wrap(~HotelType)

H1$Country[H1$Country == 'CN'] <- 'CHN' # Fixing the Country code of rows where the country code is not in ISO3 format.

library(rworldmap)
library(RColorBrewer)
par(mai=c(0,0,0.2,0),xaxs="i",yaxs="i")# The first line starting par ... below and in subsequent plots simply ensures the plot fills the available space on the page.

colourPalette <- brewer.pal(6,'YlGnBu')

nPDF <- joinCountryData2Map(Hotels, joinCode = "ISO3", nameJoinColumn = "Country")

mapCountryData( nPDF, nameColumnToPlot="Meal", colourPalette = colourPalette)
mapCountryData( nPDF, nameColumnToPlot="Meal", colourPalette = colourPalette, mapRegion = "eurasia")
mapCountryData( nPDF, nameColumnToPlot="RevenueCategory", colourPalette = colourPalette, mapRegion = "eurasia")

guest_country_pct <- Hotels%>%group_by(Country)%>%summarise(Count = n())
guest_country_pct$percent <- guest_country_pct$Count/nrow(Hotels) * 100

sPDF <- joinCountryData2Map(guest_country_pct, joinCode = "ISO3", nameJoinColumn = "Country")

mapParams <- mapCountryData(sPDF, nameColumnToPlot="percent", colourPalette = colourPalette, numCats = 5)
do.call(addMapLegend, c(mapParams
                        ,legendLabels="all"
                        ,legendWidth=0.5
                        ,legendIntervals="page"))


############################################################################
# Association rules
library(arules)
library(arulesViz)
library(ggplot2)
library(rworldmap)


Resort_hotel <- readxl::read_xlsx("H1-Resort.xlsx")
City_hotel <- readxl::read_xlsx("H2-City.xlsx")
#City_hotel<-na.omit(City_hotel)
#Resort_hotel<-na.omit(Resort_hotel)

depo <-as.factor(City_hotel$DepositType)
can<-as.factor(ifelse(City_hotel$IsCanceled==0,
                      "Not Canceled","Canceled"))
rep<-as.factor(ifelse(City_hotel$IsRepeatedGuest==0,
                      "New Guest","Repeated Guest"))
cuscateg<-as.factor(ifelse(City_hotel$Adults>2,
                           "Group","Private"))
prevcat<-as.factor(ifelse(City_hotel$PreviousCancellations>0,
                          "has cancelled before","Never cancelled before"))
prevnotcat<-as.factor(ifelse(City_hotel$PreviousBookingsNotCanceled>0,
                             "Repeat stay", "First time Stay"))
newdfCity<-data.frame(depo,cuscateg,can,rep,prevcat,prevnotcat)

depor<- as.factor(Resort_hotel$DepositType)
canr<-as.factor(ifelse(Resort_hotel$IsCanceled==0,
                       "Not Canceled","Canceled"))
repr<-as.factor(ifelse(Resort_hotel$IsRepeatedGuest==0,
                       "New Guest","Repeated Guest"))
cuscategr<-as.factor(ifelse(Resort_hotel$Adults>2,
                            "Group","Private"))
prevcatr<-as.factor(ifelse(Resort_hotel$PreviousCancellations>0,
                           "has cancelled before","Never cancelled before"))
prevnotcatr<-as.factor(ifelse(Resort_hotel$PreviousBookingsNotCanceled>0,
                              "Repeat stay", "First time Stay"))
newdfResort<-data.frame(depor,cuscategr,canr,repr,prevcatr,prevnotcatr)

Citytrans <- as(newdfCity,"transactions")
Resorttrans <-as(newdfResort,"transactions")

rulesCity <- apriori(Citytrans,parameter=list(supp=0.05, conf=0.55), control=list(verbose=F), appearance=list(default="lhs",rhs = "can=Canceled"))
summary(rulesCity)

goodrulesCity<-rulesCity[quality(rulesCity)$lift>1]
inspect(goodrulesCity[1:5])
summary(goodrulesCity)

rulesResort <- apriori(Resorttrans,
                       parameter=list(supp=0.005, conf=0.55),
                       control=list(verbose=F),
                       appearance=list(default="lhs",rhs = "canr=Canceled"))
inspect(rulesResort)
summary(rulesResort)

goodrulesResort<-rulesResort[quality(rulesResort)$lift>3]
inspect(goodrulesResort[1:5])
summary(goodrulesResort)

plot(goodrulesCity, method = "two-key plot")  
plot(goodrulesResort, method = "two-key plot") 

#########################################################
#LM models with all variables for exploring the variables that explain ADR, IsCanceled, IsRepeatedGuest 
lmADRH1 = lm(formula = ADR~ ., data = Hotels)
summary(lmADRH1)

lmH1 = lm(formula = IsCanceled~ ., data = Hotels)
summary(lmH1)

lmGuestH1 = lm(formula = IsRepeatedGuest~ ., data = Hotels)
summary(lmGuestH1)

# LM models for ADR
lmAdrH1 = lm(formula = ADR~ ReservationStatus + StaysInWeekNights + StaysInWeekendNights + IsRepeatedGuest + 
               PreviousBookingsNotCanceled + Adults + Children + ReservedRoomType + Country + DepositType + CustomerType +TotalOfSpecialRequests, data = Hotels)
summary(lmAdrH1)
## probabiblity of ADR
lmAdrH1$predict <-predict(lmAdrH1)
p = mean(lmAdrH1$predict)
p 

#Linear Model for IsRepeatedguest
lmRguestH1 = lm(formula = IsRepeatedGuest ~ LeadTime+ `Arrival Date`+ ReservationStatus + StaysInWeekendNights 
                +StaysInWeekNights + Adults + Children + Babies + Country +DistributionChannel +PreviousBookingsNotCanceled+ PreviousCancellations, data = Hotels)
summary(lmRguestH1)

# probabiblity of repeated Guest
lmRguestH1$predict <-predict(lmRguestH1) 
p = mean(lmRguestH1$predict)
p

#Linear Model for IsCanceled

lmCanceled = lm(formula = IsCanceled~ LeadTime + CustomerType + Hotel +
                  DepositType + ADR + TotalOfSpecialRequests, data = Hotels)
summary(lmCanceled)

## probabiblity of canceled the booking
lmCanceled$predict <-predict(lmCanceled) 
p = mean(lmCanceled$predict)
p 


#######################################################
#SVM
# Loading the required packages for SVM into the environment
library(kernlab) 
library(e1071) 
library(caret) 
hotel_resort <- data.frame(agent = H1$Agent, 
                           customer_type = H1$CustomerType, 
                           deposit = H1$DepositType, 
                           cancellations = as.factor(H1$PreviousCancellations), 
                           reserved_room = H1$ReservedRoomType, 
                           assigned_room = H1$AssignedRoomType, 
                           is_cancelled = as.factor(H1$IsCanceled))
str(hotel_resort)

trainList1 <- createDataPartition(y=hotel_resort$is_cancelled,p=.60,list=FALSE)
trainSet1 <- hotel_resort[trainList1,]
testSet1 <- hotel_resort[-trainList1,]

svmOut <- ksvm(is_cancelled ~ customer_type+cancellations, data=trainSet1,C=5,cross=3,prob.model=TRUE)
svmOut
table(trainSet1$is_cancelled)
table(testSet1$is_cancelled)
str(testSet)
svmPred <- predict(svmOut, testSet1)
confusionMatrix(svmPred, testSet1$is_cancelled)
# > svmPred <- predict(svmOut,testSet)
#Error in .local(object, ...) : test vector does not match model !
svmOut2 <- ksvm(is_cancelled ~ assigned_room+reserved_room, data=trainSet1, kernel='rbfdot',kpar='automatic',C=5,cross=3,prob.model=TRUE)
svmOut2
svmPred2 <- predict(svmOut2, testSet1[,-1])
table(svmPred2)
hotel_city <- data.frame(agent = as.factor(H2$Agent), 
                         customer_type = as.factor(H2$CustomerType), 
                         deposit = as.factor(H2$DepositType), 
                         cancellations = as.factor(H2$PreviousCancellations), 
                         reserved_room = as.factor(H2$ReservedRoomType), 
                         assigned_room = as.factor(H2$AssignedRoomType), 
                         is_cancelled = as.factor(H2$IsCanceled))
trainList2 <- createDataPartition(y=hotel_city$is_cancelled,p=.60,list=FALSE)
trainSet2 <- hotel_city[trainList2,]
testSet2 <- hotel_city[-trainList2,]
svmOP <- ksvm(is_cancelled ~., data=trainSet2, kernel='rbfdot',kpar='automatic',C=5,cross=3,prob.model=TRUE)
svmOP



plot(svmOut, testSet1)
length(svmOut)
#26707 observations of 31 variables 
table(trainSet$IsCanceled) #table functions give the number of classifications for each classification. Bad=120 and Good = 280
#same proportion as that of subCredit.

#We are subsetting into subCredit data frame and taking the cases which are not in trainList (-trainList) and storing this data in testSet.

svmM <- ksvm(is_cancelled~ ., data=trainSet1, C=5, cross=3, prob.model=TRUE) #Creating a svm model using ksvm() function.
svmM
#The cross validation error is 0.277 which means that about 27% of the instances that the model was learning on was mistaken.
#The cross validation error indicates that the model is not that good.

predOut <- predict(svmM, newdata=testSet, type="response") #Predicting the class of data in testSet to validate the model generated using predict() function and passing the model and data as parameters. Storing these prediction in predOut variable.
predOut #Printing the preOut variable to the console.

table(predOut)
table(predOut, testSet$IsCanceled)

# Conclusion:
### We should write something for revenue and cancellation analysis.

#Average week and weekend ARS
avg_ARS_week <- sum(H1$ARS[H1$StaysInWeekNights!=0])/NROW(H1$StaysInWeekNights[H1$StaysInWeekNights !=0])
avg_ARS_weekend <- sum(H1$ARS[H1$StaysInWeekendNights!=0])/NROW(H1$StaysInWeekendNights[H1$StaysInWeekendNights !=0])

avg_ARS_weekend / avg_ARS_week #1.24958
#Average ARS generated on weekends is greater than that of week days by 1.25 times.






