library(openxlsx)
library(readxl)
library(tidyverse)
library(broom)
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


#write.xlsx(pero %>% filter(STOCK == species) %>% slice(1:50), "test_pero.xlsx")
#write.xlsx(matingcage %>% filter(STOCK == species) %>% slice(1:25), "test_matingcage.xlsx")

#all_stock= c("BW", "LL", "PO", "IS", "EP", "SM2")

all_stock = c("SM2")






for ( species in all_stock) {
  
  
  ########################################################
  
  IND = pero %>% filter(STOCK == species)
  
  
  DAMSIRE = matingcage %>% filter(STOCK == species)
  
  
  IND2 = IND %>%
    mutate(MatingNumber = str_remove(MatingNumber, species)) %>%
    mutate(ID = str_remove(ID, species)) %>% 
    
    mutate(MatingNumber = str_replace_all(MatingNumber, "[^[:alnum:]]", "")) %>%
    
    mutate(ID = str_replace_all(ID, "[^[:alnum:]]", "")) 
  
  
  
  DAMSIRE2 = DAMSIRE %>%
    
    mutate(Dam = str_remove(Dam, species)) %>%
    mutate(Sire = str_remove(Sire, species)) %>%
    mutate(MatingNumber = str_remove(MatingNumber, species)) %>%
    mutate(Dam = str_replace_all(Dam, "[^[:alnum:]]", "")) %>%
    mutate(Sire = str_replace_all(Sire, "[^[:alnum:]]", "")) %>%
    mutate(MatingNumber = str_replace_all(MatingNumber, "[^[:alnum:]]", "")) 
  
  # Identify common columns (excluding 'MatingNumber')
  common_cols <- intersect(names(DAMSIRE2), names(IND2))
  common_cols <- setdiff(common_cols, "MatingNumber")
  
  # Remove duplicate columns from the second data frame (assuming df_pero)
  DAMSIRE2 <- DAMSIRE2 %>% select(-all_of(common_cols))
  
  
  merged_df = merge(DAMSIRE2, IND2, by = 'MatingNumber')  %>%
    mutate(Birthday = as.Date(Birthday, format = "%Y-%m-%d"), # Adjust format as needed
           BirthMonth = month(Birthday),
           BirthYear = year(Birthday))
  
  
  dam_info <- merged_df %>%
    select(ID, Birthday, BirthMonth, BirthYear) %>%
    rename(Birthday_Dam = Birthday, BirthMonth_Dam = BirthMonth, BirthYear_Dam = BirthYear)
  
  
  sire_info <- merged_df %>%
    select(ID, Birthday, BirthMonth, BirthYear) %>%
    rename(Birthday_Sire = Birthday, BirthMonth_Sire = BirthMonth, BirthYear_Sire = BirthYear)
  
  
  merged_df2 <- merged_df %>%  
    left_join(dam_info, by = c("Dam" = "ID")) %>%
    left_join(sire_info, by = c("Sire" = "ID")) 
  
  
  #############
  
  # Function to filter dataset based on Dam and Sire birth months, then count and pivot
  prepare_dataset <- function(data, dam_months, sire_months, min_year, max_year) {
    data %>%
      filter(BirthMonth_Dam %in% dam_months & BirthMonth_Sire %in% sire_months) %>%
      group_by(BirthYear, BirthMonth) %>%
      summarise(Count = n(), .groups = 'drop') %>%
      pivot_wider(names_from = BirthYear, values_from = Count, 
                  values_fill = list(Count = 0), names_prefix = "Y") %>%
      # Ensure all years are included
      complete(BirthMonth = 1:12, fill = list(Count = 0)) %>%
      # Add missing years as columns with 0
      pivot_longer(cols = starts_with("Y"), names_to = "Year", values_to = "Count") %>%
      mutate(Year = as.numeric(sub("Y", "", Year))) %>% # Convert Year back to numeric
      filter(Year >= min_year & Year <= max_year) %>%
      pivot_wider(names_from = Year, values_from = Count, values_fill = list(Count = 0))
  }
  
  # Define winter and summer months
  winter_months <- c(10, 11, 12, 1, 2, 3)
  summer_months <- setdiff(1:12, winter_months)
  
  # Calculate the overall min and max BirthYear
  min_year <- min(merged_df2$BirthYear, na.rm = TRUE)
  max_year <- max(merged_df2$BirthYear, na.rm = TRUE)
  
  Dam_winter_Sire_winter <- prepare_dataset(merged_df2, winter_months, winter_months, min_year, max_year)
  Dam_summer_Sire_summer <- prepare_dataset(merged_df2, summer_months, summer_months, min_year, max_year)
  Dam_winter_Sire_summer <- prepare_dataset(merged_df2, winter_months, summer_months, min_year, max_year)
  Dam_summer_Sire_winter <- prepare_dataset(merged_df2, summer_months, winter_months, min_year, max_year)
  
  
  
  # Create a new workbook for the current species
  wb <- createWorkbook()
  
  # Add sheets for each dataset combination and write the data
  # Directly refer to your datasets since their names don't change per species
  
  addWorksheet(wb, "Dam_winter_Sire_winter")
  writeData(wb, "Dam_winter_Sire_winter", Dam_winter_Sire_winter)  # Directly use the dataset
  
  addWorksheet(wb, "Dam_summer_Sire_summer")
  writeData(wb, "Dam_summer_Sire_summer", Dam_summer_Sire_summer)  # Directly use the dataset
  
  addWorksheet(wb, "Dam_winter_Sire_summer")
  writeData(wb, "Dam_winter_Sire_summer", Dam_winter_Sire_summer)  # Directly use the dataset
  
  addWorksheet(wb, "Dam_summer_Sire_winter")
  writeData(wb, "Dam_summer_Sire_winter", Dam_summer_Sire_winter)  # Directly use the dataset
  
  # Save the workbook with a species-specific filename
  saveWorkbook(wb, paste0("Dam_Sire_Combinations_", species, ".xlsx"), overwrite = TRUE)
  
  
  #############################
  calculate_seasonal_births_adjusted <- function(data, winter_months, summer_months, year_interval) {
    data_long <- data %>%
      pivot_longer(cols = -BirthMonth, names_to = "Year", values_to = "Count", 
                   names_prefix = "Y", names_transform = list(Year = as.integer)) %>%
      filter(Count > 0) # Keep rows with actual counts

      year_intervals <- seq(from = min_year+2, to = max_year + year_interval, by = year_interval)
    
    # Generate labels for each interval
    interval_labels <- sapply(1:(length(year_intervals)-1), function(i) {
      paste(year_intervals[i], year_intervals[i+1]-1, sep = "-")
    })
    
    # Assign records to interval groups and determine the BirthSeason
    data_long <- data_long %>%
      mutate(
        YearGroup = cut(Year, breaks = year_intervals, include.lowest = TRUE, labels = interval_labels),
        BirthSeason = case_when(
          BirthMonth %in% winter_months ~ "Winter",
          BirthMonth %in% summer_months ~ "Summer",
          TRUE ~ "Other"
        )
      )
    
    # Group by YearGroup and BirthSeason, then count
    counts <- data_long %>%
      group_by(YearGroup, BirthSeason) %>%
      summarise(Count = sum(Count), .groups = 'drop') %>%
      filter(BirthSeason %in% c("Winter", "Summer")) #%>%# Focus on Winter and Summer
    #pivot_wider(
    #  names_from = YearGroup,
    #  values_from = Count,
    #  values_fill = list(Count = 0) # Fill missing values with 0
    # )
    return(counts)
  }
  

  # Assume combinations are predefined
  combinations <- c("Dam_winter_Sire_winter", "Dam_summer_Sire_summer", "Dam_winter_Sire_summer", "Dam_summer_Sire_winter")
  
  
  
  process_and_export <- function(year_interval) {
    # Create a new workbook
    wb <- createWorkbook()
    
    # Loop through each combination
    for (combination in combinations) {
      # Retrieve the data directly without appending species to the name
      # Assuming the datasets are already in your global environment
      if (!exists(combination)) {
        next  # Skip if the dataset does not exist
      }
      
      # Since the data is not species-specific in name, directly use the variable
      data_table <- get(combination)
      
      # Calculate seasonal births
      result_counts <- calculate_seasonal_births_adjusted(data_table, winter_months, summer_months, year_interval)
      
      # Pivot result_counts for Excel export
      result_counts_pivoted <- result_counts %>%
        pivot_wider(names_from = YearGroup, values_from = Count, values_fill = list(Count = 0))
      
      # Add a worksheet to the workbook with the result
      addWorksheet(wb, combination)
      writeData(wb, combination, result_counts_pivoted)
    }
    
    # Save the workbook with a species-specific filename
    file_name <- paste("Seasonal_Births_", species, "_", year_interval, "yr.xlsx")
    saveWorkbook(wb, file_name, overwrite = TRUE)
  }
  
  
  year_intervals <- c(3)
  
  for (interval in year_intervals) {
    process_and_export( interval)
  }
  
  
}



