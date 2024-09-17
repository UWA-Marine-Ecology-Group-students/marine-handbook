# install.packages(c("Microsoft365R", "readxl", "officer"))

library(Microsoft365R)
library(readxl)
library(officer)
library(tidyverse)

# Authenticate with OneDrive
# This will prompt a browser login to authenticate
# If UWA IT get back to me we can use this, otherwise we need to sync the files to your local computer :(
# one_drive <- get_business_onedrive()
# 
# # Download Excel file from OneDrive
# excel_file <- one_drive$download_file(excel_file_path, dest = tempfile(fileext = ".xlsx"))
# 
# # Read the Excel file into R
# project_data <- read_excel(excel_file)
# 
# # Download Word template from OneDrive
# word_file <- one_drive$download_file(word_template_path, dest = tempfile(fileext = ".docx"))
# 
# # Create a new Word document object based on the template
# doc <- read_docx(word_file)

# Define file paths for the Excel file and Word template on your OneDrive
excel_file_path <- "C:/Users/00097191/OneDrive - UWA/marine-honours&masters-handbook-sbs/marine-honours-masters-projects.xlsx"
word_template_path <- "C:/Users/00097191/OneDrive - UWA/marine-honours&masters-handbook-sbs/marine-project-handbook-template.docx"

project_data <- read_excel(excel_file_path)
doc <- read_docx(word_template_path)

# Define the sections in your document where you want to insert tables
sections <- list(
  "COASTAL PROCESSES", # Placeholder, replace with actual identifier for the picture or blurb in section 1
  "OCEANOGRAPHY", # Replace with the identifier for section 2
  "MARINE ECOLOGY GROUP – FISHERIES RESEARCH", # Same for section 3
  "pic4-blurb"  # Same for section 4
)

# Create a loop to add tables for each row in the Excel sheet
# Loop through the rows in the Excel sheet
for (i in 1:nrow(project_data)) {
  # Extract each row's information
  row_data <- project_data[i, ]
  
  # Create a table for this row
  table_content <- data.frame(
    Property = c("Project Title", "Supervisors", "Description", "Start Date", "Requirements"),
    Details = c(row_data$ProjectTitle, row_data$Supervisors, row_data$Description, row_data$StartDate, row_data$Requirements)
  )
  
  # Identify the section this row belongs to
  section_name <- row_data$Group
  
  # Find the section in the document where the table should be added
  # if (section_name %in% names(sections)) {
    # Use a cursor to move to the correct section in the document
    doc <- doc %>%
      cursor_reach(sections[[section_name]]) %>%  # Reach the section's picture/blurb
      body_add_par("", style = "Normal") %>%      # Add a new paragraph after the heading
      body_add_table(table_content, style = "table_grid") %>%  # Insert table after the picture/blurb
      body_add_par(value = "", style = "Normal")  # Adding space between tables
  # }
}

# Load the Word template
doc <- read_docx(word_template_path)

# Loop through the rows in the Excel sheet
for (i in 1:nrow(project_data)) {
  # Extract each row's information
  row_data <- project_data[i, ]
  
  # Create a table for this row
  table_content <- data.frame(
    Property = c("Project Title", "Supervisors", "Description", "Start Date", "Requirements"),
    Details = c(row_data$ProjectTitle, row_data$Supervisors, row_data$Description, row_data$StartDate, row_data$Requirements)
  )
  
  # Extract the header text from the Excel sheet (now named 'Group')
  header_text <- row_data$Group
  
  # Find the section in the document where the table should be added based on the header text
  if (!is.na(header_text) && header_text != "") {
    # Use the actual header text from the Excel file to find the section
    doc <- doc %>%
      cursor_reach(header_text) %>%      # Reach the specific header text
      body_add_par("", style = "Normal") %>%      # Add a new paragraph after the heading
      body_add_table(table_content, style = "table_grid") %>%  # Insert the table after the header
      body_add_par("", style = "Normal")          # Add space after the table
  }
}


doc <- doc %>%
  cursor_reach(section_name)

# # Upload the modified document back to OneDrive
# one_drive$upload_file(output_file, dest = word_template_path, overwrite = TRUE)

# Save the modified Word document locally
local_output_file <- "C:/Users/00097191/OneDrive - UWA/marine-honours&masters-handbook-sbs/marine-project-handbook.docx"
print(doc, target = local_output_file)

cat("The Word document has been updated successfully and uploaded to OneDrive.\n")

