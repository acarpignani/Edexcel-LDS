getwd()
setwd("~/Documents/Edexcel Large Data Set/")

library(tidyverse)
library(readxl)
library(janitor)

all_sheets <- excel_sheets("./Pearson Edexcel GCE AS and AL Mathematics data set.xls")
sheets_uk <- all_sheets[2:11]
sheets_overseas <- all_sheets[-(1:11)]

#
# UK Stations
#
workbook_uk <- lapply(sheets_uk, 
                   read_excel, 
                   path = "./Pearson Edexcel GCE AS and AL Mathematics data set.xls",
                   col_types = c("date", rep("guess", 14)),
                   na = c("na", "NA", "n/a", "#N/A", "tr"),
                   skip = 5)
names(workbook_uk) <- sheets_uk

pearson_uk <- lapply(sheets_uk,
       \(.) workbook_uk[[.]] |> 
           rename(date = 1, temperature = 2, rainfall = 3, sunshine = 4, 
                  windspeed = 5, wind_beaufort = 6, gust = 7, 
                  humidity = 8, cloud = 9, visibility = 10, pressure = 11, 
                  wind_direction = 12, wind_cardinal = 13, 
                  gust_direction = 14, gust_cardinal = 15) |> 
           mutate_at(vars(wind_beaufort,
                          wind_cardinal,
                          gust_cardinal), 
                     factor) |> 
           mutate(date = as.Date(date), station = word(., 1), .after = date)
           ) |> 
    bind_rows()

#
# Overseas Stations
#
workbook_overseas <- lapply(sheets_overseas, 
                      read_excel, 
                      path = "./Pearson Edexcel GCE AS and AL Mathematics data set.xls",
                      col_types = c("date", rep("guess", 5)),
                      na = c("na", "NA", "n/a", "#N/A", "tr"),
                      skip = 5)
names(workbook_overseas) <- sheets_overseas

pearson_overseas <- lapply(sheets_overseas,
                           \(.) workbook_overseas[[.]] |> 
                               rename(date = 1, temperature = 2, rainfall = 3, 
                                      pressure = 4, windspeed = 5, 
                                      wind_beaufort = 6) |> 
                               mutate(
                                   wind_beaufort = factor(wind_beaufort),
                                   date = as.Date(date), 
                                   station = word(., 1), .after = date)
) |> 
    bind_rows()

pearson = list(uk = pearson_uk, overseas = pearson_overseas)

# Saving result
write_csv(pearson_uk, "./data/pearson_uk.csv")
write_csv(pearson_overseas, "./data/pearson_overseas.csv")
save(pearson, file = "./data/pearson.RData")
