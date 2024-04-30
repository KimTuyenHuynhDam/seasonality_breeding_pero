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


# List of species
all_stock <- c("BW", "LL", "PO", "IS", "EP", "SM2")

# Initialize a workbook and add a worksheet
wb <- createWorkbook()
sheet_name <- "Species_Summary"
addWorksheet(wb, sheet_name)

# Initialize a data frame to hold all the summary data
summary_df <- data.frame(Species = character(), 
                         Period = character(), 
                         TotalBirths = integer(), 
                         TotalPairs = integer(), 
                         TotalOffspring = integer())


# Processing for each species
for (species in all_stock) {
  
  # Filter data for the species
  
  IND = pero %>% filter(STOCK == species)
  
  
  DAMSIREori = matingcage %>% filter(STOCK == species)
  
  # Filter DAMSIRE to include only rows where MatingNumber exists in IND
  DAMSIRE <- DAMSIREori %>%
    semi_join(IND, by = "MatingNumber")
  
  
  # Define period range based on species
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
  
  # Period description
  period_description <- paste(start_year, "to", end_year)
  
  # Append the data for this species to the summary dataframe
  summary_df <- rbind(summary_df, data.frame(
    Species = species,
    Period = period_description,
    TotalBirths = total_births,
    TotalPairs = total_pairs,
    TotalOffspring = total_offspring
  ))
}

# Write the complete data frame to the Excel sheet
writeData(wb, sheet_name, summary_df)

# Save the workbook
saveWorkbook(wb, "Total_Species_Data_Summary.xlsx", overwrite = TRUE)