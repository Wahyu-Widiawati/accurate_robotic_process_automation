# Set working directory
setwd("C:/Users/User/Documents/Cedea/SI Accurate")

#Install Pacakages
if(!require(googledrive))install.packages("googledrive");library(googledrive)
if(!require(googlesheets4))install.packages("googlesheets4");library(googlesheets4)
if(!require(readxl))install.packages("readxl");library(readxl)
#if(!require(readxlsb))install.packages("readxlsb");library(readxlsb)
if(!require(RODBC)) install.packages('RODBC'); require(RODBC)
if(!require(tidyverse))install.packages("tidyverse");library(tidyverse)
if(!require(lubridate))install.packages("lubridate");library(lubridate)
if(!require(DBI))install.packages("DBI");library(DBI) #to execute SQL query
if(!require(RPostgres))ianstall.packages("RPostgres");library(RPostgres)#to connect RPostgres
if(!require(lubridate))install.packages("lubridate");library(lubridate)#to deal with date format
if(!require(zoo))install.packages("zoo");library(zoo)
if(!require(tools))install.packages("tools");library(tools)

# Connect to the Database
con <- dbConnect(
  RPostgres::Postgres(),
  dbname = "yyy",
  host = "xxx",
  port = "5432",
  user = "cda_ds_team",
  password = "zzz"
)

# list file in local
list_file <- list.files(path= "C:/Users/User/Documents/Cedea/SI Accurate")
dfAll <- list()
dfLeft <- list()

# Define your column names once
column_names <- c(
  'proceed_type','invoice_no', 'invoice_date', 'warehouse', 'item_no', 'product_description',
  'product_category', 'invoiced_qty','unit', 'amount', 'cogs_amount', 'customer_no',
  'customer_name', 'invoice_ship','channel', 'order_no', 'order_date',
  'inclusive_tax', 'discount', 'discount_global', 'disc_line', 'invoice_description', 
  'ship_date', 'invoice_status', 'unit_price', 'tax_amount', 'sequence', 'total_price'
)

for(file_name in list_file){
  file_exte <- tools::file_ext(file_name)
  if (file_exte == 'xls'){
    file_path <- paste0("C:/Users/User/Documents/Cedea/SI Accurate/", file_name)
    df <- read_xls(file_path)
    
    if(ncol(df) > 2){
      df <- df[, colSums(!is.na(df)) > 0]  # Remove all-NA columns
      
      # Locate the "Invoice" header row
      header_row_index <- which(df[[1]] == "Invoice")[1]
      
      if (!is.na(header_row_index) && nrow(df) > header_row_index) {
        # Slice data after the header
        df_clean <- df[(header_row_index):nrow(df), ]
        
        # Apply column names to this df
        colnames(df_clean) <- column_names
        
        # Remove rows where proceed_type is not "Invoice"
        df_clean <- df_clean %>% filter(proceed_type == 'Invoice')
        
        # Save cleaned df to list
        dfAll[[file_name]] <- df_clean
      } else {
        cat("Skipping file (no valid data found after header):", file_name, "\n")
        dfLeft[[length(dfLeft) + 1]] <- file_name
      }
    } else {
      cat("Skipping file (not enough columns):", file_name, "\n")
      dfLeft[[length(dfLeft) + 1]] <- file_name
    }
  }
}

# Combine all data frames
dfAllSO <- data.frame()

# Combine them into a single data frame
dfAllSO <- do.call(rbind, dfAll)

# # Change the data type
dfAllSO$invoiced_qty <- as.double(gsub(",", ".", dfAllSO$invoiced_qty))
dfAllSO$unit_price <- as.double(gsub(",", ".", dfAllSO$unit_price))
dfAllSO$discount <- as.double(gsub(",", ".", dfAllSO$discount))
dfAllSO$total_price <- as.double(gsub(",", ".", dfAllSO$total_price))
dfAllSO$amount <- as.double(gsub(",", ".", dfAllSO$amount))
dfAllSO$tax_amount <- as.double(gsub(",", ".", dfAllSO$tax_amount))
dfAllSO$invoice_date <- as.Date(dfAllSO$invoice_date, format="%d %b %Y")
dfAllSO$order_date <- as.character(dfAllSO$order_date)
dfAllSO$order_date <- as.Date(dfAllSO$order_date, format="%d %b %y")
dfAllSO$ship_date <- as.character(dfAllSO$ship_date)
dfAllSO$ship_date <- as.Date(dfAllSO$ship_date, format="%d %b %y")

# Add a new Column to Calculate Total Price - Discount Line
dfAllSO$disc_line <- as.numeric(dfAllSO$disc_line)
dfAllSO$price_subtotal <- NA # total price - line discount
dfAllSO <- dfAllSO %>% 
  mutate(price_subtotal = ifelse(is.na(disc_line), 
                                 invoiced_qty * unit_price, 
                                 (100 - disc_line) / 100 * invoiced_qty * unit_price))

# Add a new column to calculate the subtotal
dfAllSO <- dfAllSO %>%
  group_by(invoice_no) %>%
  mutate(subtotal = sum(price_subtotal, na.rm = TRUE)) %>%
  ungroup()

DBI::dbExecute(con, "DELETE FROM cda_it_custom.sale_order_line_accurate_history")

dbWriteTable(con, 
             Id(schema = "cda_it_custom", table = "sale_order_line_accurate_history"), 
             dfAllSO, 
             append = TRUE)

