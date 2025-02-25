library(openxlsx)
library(readxl)
library(dplyr)
library(broom)
library(tidyverse)
library(lubridate)

# Function to replace spaces with underscores in column names
clean_column_names <- function(df) {
  names(df) <- gsub(" ", "", names(df))
  return(df)
}

pero = read_excel("Peromyscus.xlsx") %>% select(1:5)  %>%
  clean_column_names %>% filter(!is.na(Birthday)) 
matingcage  = read_excel("Mating Records.xlsx")  %>% select(1:5) %>%
  clean_column_names



all_stock <- c("BW", "LL", "PO", "IS", "EP", "SM2")


wb <- createWorkbook()
sheet_name <- "Species_Summary"
addWorksheet(wb, sheet_name)


summary_df <- data.frame(Species = character(), 
                         Period = character(), 
                         TotalBirths = integer(), 
                         TotalPairs = integer(), 
                         TotalOffspring = integer())


# Processing for each species
for (species in all_stock) {
  

  IND = pero %>% filter(STOCK == species)
  
  
  DAMSIREori = matingcage %>% filter(STOCK == species)
 
  DAMSIRE <- DAMSIREori %>%
    semi_join(IND, by = "MatingNumber")
  
  
 
  if (species == "SM2") {
    start_year <- min(year(IND$Birthday) +2, na.rm = TRUE)
    end_year <- 2016
  } else {
    start_year <- min(year(IND$Birthday)+1, na.rm = TRUE) + 1
    end_year <- 2023
  }
  
  
  # Calculate total births for the period
  
  total_births <- IND %>%
    filter(year(Birthday) >= start_year & year(Birthday) <= end_year) %>%
    summarise(TotalBirths = n_distinct(Birthday)) %>% # Count unique birthdays
    pull(TotalBirths)
  
  # Calculate total breeding pairs for the period
  total_pairs <- DAMSIRE %>%
    filter(year(DateofMating) >= start_year & year(DateofMating) <= end_year) %>%
    summarise(TotalPairs = n()) %>%
    pull(TotalPairs)
  
  # Calculate total offspring for the period
  total_offspring <- IND %>%
    filter(year(Birthday) >= start_year & year(Birthday) <= end_year) %>%
    summarise(TotalOffspring = n()) %>%
    pull(TotalOffspring)
  
 
  period_description <- paste(start_year, "to", end_year)
  
  
    Species = species,
    Period = period_description,
    TotalBirths = total_births,
    TotalPairs = total_pairs,
    TotalOffspring = total_offspring
  ))
}


writeData(wb, sheet_name, summary_df)


saveWorkbook(wb, "Total_Species_Data_Summary.xlsx", overwrite = TRUE)
