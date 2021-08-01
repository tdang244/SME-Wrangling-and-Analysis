# Install and load libraries: Remember to run these first
install.packages("magrittr") 
install.packages("dplyr")
install.packages("data.table")
install.packages("writexl")

library(magrittr)
library(dplyr)
library(data.table)
library(readxl)
library(writexl)

# Upload data tables: Keep the names of the data tables. I'll enclose the whole folder with the data for you. Keep this file in the same folder with the data tables. 

all_data <- read_excel("all_data.xlsx")
province_code <- read_excel("province_code.xlsx")
industry_info <- read_excel("industry_info.xlsx")
hdb_sme_data <- read_excel("Client_SME.xlsx")
branch_data <- read_excel("branch_data.xlsx", 
                          sheet = "support 1")


# Describe the data
summary(all_data)

# Clean province code
cleaner_province_code <- unique(province_code)

# Set conditions to filter out micro se, se, me, and mmlc
all_data_copy <- copy(unique(all_data))
all_data_copy$segment <- case_when(
  all_data_copy$revenue <= 9200 ~ "micro se",
  all_data_copy$revenue > 9200 & all_data_copy$revenue <= 114000 ~ "se",
  all_data_copy$revenue > 114000 & all_data_copy$revenue <= 460000 ~ "me",
  all_data_copy$revenue > 460000 ~ "mmlc",
  TRUE ~ as.character(all_data_copy$revenue)
)
str(all_data_copy)

# Count Micro SEs, SEs, MEs, and Large corps 
segment_count <- all_data_copy %>% group_by(segment) %>% summarise(company_type_count = n_distinct(name))

# Filter out SEs
se_data_only <- all_data_copy %>% filter(segment == "se")

# Join SE data with industry data
industry_se_data <- setDT(merge(se_data_only, industry_info, by.x = "industry_code", by.y = "level5", all.x = TRUE))

# Join the data right above with the province data
province_industry_se_data <- setDT(merge(industry_se_data, cleaner_province_code, by.x = "province", by.y = "code", all.x = TRUE))

# Choose relevant columns for the final SE table
clean_se_data <- province_industry_se_data[, c("province.y", "overall_industry_name", "sub_industry_name", "name.x", "asset", "revenue", "segment", "industry_code", "level1")]

setnames(clean_se_data, c("province.y", "name.x", "level1"), c("province", "company_name", "overall_industry_code"))

# clean_se_data %>% distinct(segment) # checking segment count

# Final SE table
se_final <- clean_se_data %>% group_by(province, overall_industry_name) %>% summarise(se_count = n_distinct(company_name), se_total_revenue = sum(revenue), se_total_asset_size = sum(asset))

# Filter out MEs
me_data_only <- all_data_copy %>% filter(segment == "me")

# Join ME data with industry data
industry_me_data <- setDT(merge(me_data_only, industry_info, by.x = "industry_code", by.y = "level5", all.x = TRUE))

# Join the data right above with province data
province_industry_me_data <- setDT(merge(industry_me_data, cleaner_province_code, by.x = "province", by.y = "code", all.x = TRUE))

# Choose relevant columns for ME data
clean_me_data <- province_industry_me_data[, c("province.y", "overall_industry_name", "sub_industry_name", "name.x", "asset", "revenue", "segment", "industry_code", "level1")]

setnames(clean_me_data, c("province.y", "name.x", "level1"), c("province", "company_name", "overall_industry_code"))

# clean_me_data %>% distinct(segment) # checking segment count

# Final ME data
me_final <- clean_me_data %>% group_by(province, overall_industry_name) %>% summarise(me_count = n_distinct(company_name), me_total_revenue = sum(revenue), me_total_asset_size = sum(asset))

# Full table with SME count, revenue total, and asset total
sme_rev_asset <- setDT(merge(se_final, me_final, by.x = c("province", "overall_industry_name"), by.y = c("province", "overall_industry_name"), all = TRUE))

# Join the preliminary SME table with HDB SME data
sme_add_hdb_info <- setDT(merge(sme_rev_asset, hdb_sme_data, by.x = c("province", "overall_industry_name"), by.y = c("province", "industry"), all = TRUE))

sme_add_hdb_info <- sme_add_hdb_info %>% filter(is.na(province) == FALSE)

# Join the table right above with the data on HDB branch and total SME by province
final_table <- merge(sme_add_hdb_info, branch_data, by = "province", all = TRUE)
clean_final_table <- final_table[, !c("no", "location")]
setnames(clean_final_table, "hdb_count", "hdb_office_count")

sorted_final_table <- clean_final_table[order(province)]

# Export the table to excel. You will find this table in the same folder with your other datasets. 
write_xlsx(sorted_final_table,"Result - SME Geographical Analysis.xlsx")

# Count of SE and ME by industry only
se_by_ind <- clean_se_data %>% group_by(overall_industry_name) %>% summarise(se_count = n_distinct(company_name), se_total_revenue = sum(revenue), se_total_asset_size = sum(asset))

me_by_ind <- clean_me_data %>% group_by(overall_industry_name) %>% summarise(me_count = n_distinct(company_name), me_total_revenue = sum(revenue), me_total_asset_size = sum(asset))

se_me_by_ind <- setDT(merge(se_by_ind, me_by_ind, by.x = c("overall_industry_name"), by.y = c("overall_industry_name"), all = TRUE))

# Export the SME count by industry into excel
write_xlsx(se_me_by_ind,"SME Count by Industry.xlsx")

