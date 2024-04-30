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

all_stock= c("BW", "LL", "PO", "IS", "EP", "SM2")

all_stock = c("SM2")
#species = c("BW")





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
  
  prepare_dataset <- function(data, min_year, max_year) {
    # Initial grouping and summarizing
    summarized_data <- data %>%
      group_by(BirthYear, BirthMonth) %>%
      summarise(Count = n(), .groups = 'drop')
    
    # Creating a sequence of years to ensure all years are represented
    all_years <- paste0("Y", seq(min_year, max_year))
    
    # Pivoting to wide format, ensuring all years are included
    wide_data <- summarized_data %>%
      pivot_wider(names_from = BirthYear, values_from = Count, 
                  values_fill = list(Count = 0), names_prefix = "Y") %>%
      complete(BirthMonth = 1:12, fill = list(Count = 0))
    
    # Ensure columns for all years in the range, adding missing ones as necessary
    missing_years <- setdiff(all_years, names(wide_data))
    for(year in missing_years) {
      wide_data[[year]] <- 0
    }
    
    # Reorder columns to maintain the chronological order of years
    wide_data <- wide_data %>%
      select(BirthMonth, sort(names(wide_data)[-1]))
    
    return(wide_data)
  }
  
  # Calculate the overall min and max BirthYear from merged_df2
  min_year <- min(merged_df2$BirthYear, na.rm = TRUE) +2
  max_year <- max(merged_df2$BirthYear, na.rm = TRUE)
  
  # Prepare the dataset
  all_mice <- prepare_dataset(merged_df2, min_year, max_year)
  
  
  # Create a new workbook
  wb <- createWorkbook()
  
  # Add a worksheet to the workbook
  addWorksheet(wb, "All Mice")
  
  # Write your data to the worksheet
  writeData(wb, "All Mice", all_mice)
  
  # Define the file name with species-specific information
  file_name <- paste0("all_F1_mice born from ", min_year, " to ", max_year, " - ", species, ".xlsx")
  
  # Save the workbook
  saveWorkbook(wb, file_name, overwrite = TRUE)
  
  winter_months <- c(10, 11, 12, 1, 2, 3)
  summer_months <- setdiff(1:12, winter_months)
  
  
  
  #####################
  calculate_seasonal_births <- function(data, winter_months, summer_months, year_interval) {
    data_long <- data %>%
      pivot_longer(cols = -BirthMonth, names_to = "Year", values_to = "Count", 
                   names_prefix = "Y", names_transform = list(Year = as.integer)) %>%
      filter(Count > 0) # Keep rows with actual counts
    
    # Determine the year range for intervals based on the specified year_interval
    #min_year <- min(data_long$Year, na.rm = TRUE)
    #ax_year <- max(data_long$Year, na.rm = TRUE)
    year_intervals <- seq(from = min_year, to = max_year + year_interval, by = year_interval)
    
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
  ###########################
  process_and_export <- function(data, winter_months, summer_months, year_intervals) {
    wb <- createWorkbook()
    
    for (year_interval in year_intervals) {
      result_counts <- calculate_seasonal_births(data, winter_months, summer_months, year_interval)
      
      result_counts_pivoted <- result_counts %>%
        pivot_wider(names_from = YearGroup, values_from = Count, values_fill = list(Count = 0))
      
      sheet_name <- paste("Interval", year_interval, "yr")  # Use year_interval to name the sheet
      addWorksheet(wb, sheet_name)
      writeData(wb, sheet_name, result_counts_pivoted)
    }
    
    
    file_name <- paste("Seasonal_Births_", species, "_", paste(year_intervals, collapse = "_"), "yr.xlsx")
    saveWorkbook(wb, file_name, overwrite = TRUE)
  }
  
  
  year_intervals <- c(3, 4, 5)
  for (interval in year_intervals) {
    process_and_export(all_mice, winter_months, summer_months, year_intervals)
  }
}