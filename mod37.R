## Code for MOD 37 Part I

fleetplan <- read.csv(file = "fleetplan.csv", header = F)

# Separate data, and get rid of possible empty cells resulting from the excel manipulations
library(dplyr)
library(openxlsx)

datta <- fleetplan %>% select(V3:V6, V8)
datta <- na.omit(datta)
buscost <- fleetplan %>% select(V9:V12)
buscost <- na.omit(buscost)
mileage <- fleetplan %>% select(V13, V14)

# Rename the columns
names(datta) <- c("equip.code", "part.price", "hours", "description", "equip.type")
names(mileage) <- c("equip.code", "mile")

# Remove the bars from the mileage data
mileage$equip.code <- as.numeric(sub('.|', '', mileage$equip.code))
mileage$mile <- as.numeric(sub('*\\|', '', mileage$mile))

# Scheduled data
library(stringr)
sdata <- datta %>% filter(str_detect(description,  "Engine Replac")|str_detect(description,  "3K PM")|
                            str_detect(description,  "Turbo Replac")|str_detect(description, "engine replac")|
                            str_detect(description,  "Transmission Replace")|str_detect(description,"Transmission Replac")|
                            str_detect(description,  "turbo Replaceme")|str_detect(description,  "3 K PM"))
Labor.Total <- 38.84 * sum(sdata$hours)
Parts.Total <- sum(sdata$part.price)                          
Scheduled.Total <- Labor.Total + Parts.Total
Scheduled.Data <- sdata %>% summarise(Labor.Total, Parts.Total, Scheduled.Total)

# Summary

# Attach miles information
names(buscost) <- c("a", "b", "c", "series")
bus.series.mile <- as.data.frame(cbind(buscost$series, na.omit(mileage)))

# compute numbers
names(bus.series.mile) <- c("equip.type", "equip.code", "mile") # renaming the mile column

my_table <- aggregate(mile ~ equip.type, data = bus.series.mile, sum) 
cost.per.mile <- c(0.47, 0.09, 0.05, 0.12)

# add $/mile and cost info
my_table2 <- my_table %>% mutate(cost.per.mile)
my_table3 <- my_table2 %>% mutate(cost = mile * cost.per.mile)

# retrieve and attach formerly computed numbers
Total.Mile.Cost <- as.numeric(c(sum(my_table3$cost),"N/A", "N/A", "N/A"))
Total.Scheduled.Cost <- as.numeric(c(Scheduled.Total, "N/A", "N/A", "N/A"))
Modification.Cost <- Total.Mile.Cost + Total.Scheduled.Cost
table_summary <- as.data.frame(cbind(my_table3, Total.Mile.Cost, Total.Scheduled.Cost, Modification.Cost))

## CREATE EXCEL WORKBOOK AND EXPORT THE REPORT

# create a nd name the sheet to the workbook
wbb <- createWorkbook("MOD37 Part I")
addWorksheet(wbb, "Summary Part I")
addWorksheet(wbb, "Mileage Data")
addWorksheet(wbb, "Scheduled Data")
addWorksheet(wbb, "Raw")
addWorksheet(wbb, "Bust Cost data")
addWorksheet(wbb, "Schd. Summary")
addWorksheet(wbb, "Data for Summary")


# Move the reports the sheets
writeData(wbb, sheet = 1, table_summary)
writeData(wbb, sheet = 2, mileage)
writeData(wbb, sheet = 3, sdata)
writeData(wbb, sheet = 4, datta)
writeData(wbb, sheet = 5, buscost)
writeData(wbb, sheet = 6, Scheduled.Data)
writeData(wbb, sheet = 7, bus.series.mile)

# export the filnal output to documents
saveWorkbook(wbb, "MOD37_Part_One.xlsx", overwrite = TRUE)

##  Part II

mod37 <- read.csv(file = "mod37.csv", header = F)

# Separate data, and get rid of possible empty cells resulting from the excel manipulations
datta2 <- mod37 %>% select(V3:V6, V8)
datta2 <- na.omit(datta2)
buscost2 <- mod37 %>% select(V9:V12)
buscost2 <- na.omit(buscost2)
mileage2 <- mod37 %>% select(V13, V14)
mileage2 <- na.omit(mileage2)

# Rename the columns
names(datta2) <- c("equip.code", "part.price", "hours", "description", "equip.type")
names(mileage2) <- c("equip.code", "mile")

# Remove the bars from the mileage data
mileage2$equip.code <- as.numeric(sub('.|', '', mileage2$equip.code))
mileage2$mile <- as.numeric(sub('*\\|', '', mileage2$mile))

# Scheduled data
sdata2 <- datta2 %>% filter(str_detect(description,  "Engine Replac")|str_detect(description,  "3K PM")|
                              str_detect(description,  "Turbo Replac")|str_detect(description, "engine replac")|
                              str_detect(description,  "Transmission Replace")|str_detect(description,"Transmission Replac")|
                              str_detect(description,  "turbo Replaceme")|str_detect(description,  "3 K PM"))
Labor.Total2 <- 38.84 * sum(sdata2$hours)
Parts.Total2 <- sum(sdata2$part.price)                          
Scheduled.Total2 <- Labor.Total2 + Parts.Total2
Scheduled.Data2 <- sdata2 %>% summarise(Labor.Total2, Parts.Total2, Scheduled.Total2)

# Summary

# Attach miles information
names(buscost2) <- c("a", "b", "c", "series")
bus.series.mile2 <- as.data.frame(cbind(buscost2$series, na.omit(mileage2)))

# compute numbers
names(bus.series.mile2) <- c("equip.type", "equip.code", "mile") # renaming the mile column

my_table.b <- aggregate(mile ~ equip.type, data = bus.series.mile2, sum) 
cost.per.mile.b <- c(0.05, 0.12, 0.06, 0.39)

# add $/mile and cost info
my_table2.b <- my_table.b %>% mutate(cost.per.mile.b)
my_table3.b <- my_table2.b %>% mutate(cost = mile * cost.per.mile.b)

# retrieve and attach formerly computed numbers
Total.Mile.Cost2 <- as.numeric(c(sum(my_table3.b$cost),"N/A", "N/A", "N/A"))
Total.Scheduled.Cost2 <- as.numeric(c(Scheduled.Total2, "N/A", "N/A", "N/A"))
Modification.Cost2 <- Total.Mile.Cost2 + Total.Scheduled.Cost2
table_summary2 <- as.data.frame(cbind(my_table3.b, Total.Mile.Cost2, Total.Scheduled.Cost2, Modification.Cost2))

## CREATE EXCEL WORKBOOK AND EXPORT THE REPORT

# create a nd name the sheet to the workbook
wbb2 <- createWorkbook("MOD37 Part I")
addWorksheet(wbb2, "Summary Part II")
addWorksheet(wbb2, "Mileage Data")
addWorksheet(wbb2, "Scheduled Data")
addWorksheet(wbb2, "Raw")
addWorksheet(wbb2, "Bust Cost data")
addWorksheet(wbb2, "Schd. Summary")
addWorksheet(wbb2, "Data for Summary")

# Move the reports the sheets
writeData(wbb2, sheet = 1, table_summary2)
writeData(wbb2, sheet = 2, mileage2)
writeData(wbb2, sheet = 3, sdata2)
writeData(wbb2, sheet = 4, datta2)
writeData(wbb2, sheet = 5, buscost2)
writeData(wbb2, sheet = 6, Scheduled.Data2)
writeData(wbb2, sheet = 7, bus.series.mile2)

# export the final output to documents
saveWorkbook(wbb2, "MOD37_Part_Two.xlsx", overwrite = TRUE)

##  EMAIL THE REPORT OUT
library(RDCOMClient)
## Open outlook
outlook <- COMCreate("Outlook.Application")

## create new email
email <- outlook$CreateItem(0)

## Set recipients and body
email[["to"]] <- "joel.messan@ratpdev.com; paul.howell@ratpdev.com; joemessan@gmail.com" 
email[["subject"]] <- "MOD37 Reports"
email[["body"]] <- "Please see attached, the MOD 37 Part I and II.

Best regards,
Joel Messan"
email[["attachments"]]$Add("C:\\Users\\jmessan\\Documents\\MOD37_Part_One.xlsx")
email[["attachments"]]$Add("C:\\Users\\jmessan\\Documents\\MOD37_Part_Two.xlsx")

## Send email
email$Send()

## close outlook, clear message
rm(outlook, email)