library(officer)
library(readxl)
library(dplyr)

# Load the Excel data
load_data <- function(file_path) {
  read_excel(file_path)
}

# Create a PowerPoint presentation
create_presentation <- function(data) {
  # Create a new PowerPoint file
  ppt <- read_pptx()
  
  # Add a title slide
  ppt <- add_slide(ppt, layout = "Title Slide", master = "Office Theme")
  ppt <- ph_with(ppt, value = "2024 ABC Conference Agenda", location = ph_location_type(type = "title"))
  ppt <- ph_with(ppt, value = "Welcome to the Conference!", location = ph_location_type(type = "subtitle"))
  
  # Add slides for each speaker
  data %>%
    filter(!is.na(Speaker)) %>%
    mutate(Time = ifelse(is.na(Time), "", Time),
           Speaker = ifelse(is.na(Speaker), "", Speaker),
           Program_Section = ifelse(is.na(Program_Section), "", Program_Section),
           LinkedIn = ifelse(is.na(LinkedIn), "", LinkedIn),
           Title_Affiliation = ifelse(is.na(Title_Affiliation), "", Title_Affiliation)) %>%
    rowwise() %>%
    do({
      slide_data <- .
      ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
      ppt <- ph_with(ppt, value = paste(slide_data$Time, slide_data$Program_Section), location = ph_location_type(type = "title"))
      ppt <- ph_with(ppt, value = paste("Speaker:", slide_data$Speaker, 
                                       "\nTitle & Affiliation:", slide_data$Title_Affiliation,
                                       "\nDuration:", slide_data$Min, "minutes",
                                       "\nLinkedIn:", slide_data$LinkedIn), 
                     location = ph_location_type(type = "content"))
    })
  
  # Save the presentation
  print(ppt, target = "Conference_Agenda_Presentation.pptx")
  cat("Presentation created successfully!\n")
}

# Main function to run the tasks
main <- function() {
  file_path <- "2024 ABC Agenda.xlsx"
  data <- load_data(file_path)
  create_presentation(data)
}

# Execute the main function
main()
