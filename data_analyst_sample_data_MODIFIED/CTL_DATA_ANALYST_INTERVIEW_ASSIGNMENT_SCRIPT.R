###############################################################################
## File Name: CTL_Data_Analyst_Interview_R_Script.R
## Name: Nicole Segaran
## Script purpose: This script takes in the given excel event attendance files 
## and outputs a graph of the number of attending students per Stanford school
###############################################################################

library(readxl)
library(tidyverse)

## Function: for_each_wkbk
## This function takes in a path to one of the given excel files and the most
## updated data frame of registrees, their schools, and attendees.  It returns
## the data frame updated with information from the current excel file.
###############################################################################
for_each_wkbk <- function(export_data_old, excel_path) {
  registered <- read_excel(excel_path, sheet = 1)  # Import and reformat an excel workbook
  attended <- read_excel(excel_path, sheet = 2)
  names(registered)<-str_replace_all(names(registered), c(" " = ".")) #clean up column names to acceptable formats
  names(attended)<-str_replace_all(names(attended), c(" " = ".", "[(]" = "", "[)]" = ""))
  registered$registrees <- paste(registered$First.Name, registered$Last.Name)
  max_length <- max(c(length(registered$registrees), length(attended$Name.Original.Name)))  #Update data frame with relevant data from current workbook
  export_data_new <- data.frame(attendees_name = c(attended$Name.Original.Name, rep(NA, max_length - length(attended$Name.Original.Name))),
                                registrees_name = c(registered$registrees,rep(NA, max_length - length(registered$registrees))),
                                registrees_school = c(registered$Stanford.School, rep(NA, max_length - length(registered$Stanford.School))))
  export_data_updated <- rbind(export_data_old, export_data_new)
  return(export_data_updated)
}


## Function: clean_data_frame
## This function takes in a version of the data frame updated with all excel
## file information, and cleans it so that it may be used to create a bar plot
## (i.e. deleting duplicates, registrees who did not attend an event, etc.).
## It returns a final, cleaned data frame.
###############################################################################
clean_data_frame <- function(data_frame) {
  data_frame <- data_frame %>% 
    mutate(attendees_school= if_else(data_frame$registrees_name %in% data_frame$attendees_name, data_frame$registrees_school, "no show"))
  data_frame <- data_frame %>% 
    filter(duplicated(data_frame$registrees_name) == FALSE)
  data_frame_cleaned <- data_frame[!(data_frame$attendees_school == "no show"), ]
  return(data_frame_cleaned)
}

## This section executes entire program to output data visualization
###############################################################################

setwd("Desktop/data_analyst_sample_data_MODIFIED") # set working directory
files.list <- list.files(pattern='*.xlsx', recursive = TRUE) #create list of all excel files in directory
export_data_old <- data.frame(attendees_name = c(NA), registrees_name = c(NA), registrees_school = c(NA)) #Create initial data frame
for (i in 1:length(files.list)){  #loop through excel workbooks
  excel_path <- files.list[i]
  export_data_old <- for_each_wkbk(export_data_old, excel_path)
}
pdf(file = 'Students Per School Graph.pdf') #Create and export data visualization
ggplot(clean_data_frame(export_data_old), aes(x = attendees_school, fill = attendees_school)) + geom_bar() + ggtitle("Number of Students Per School") + xlab("Stanford Schools") + ylab("Number of Students") + theme(plot.title = element_text(size=14, face="bold"), axis.title.x = element_text(size=14, face="bold"), axis.title.y = element_text(size=14, face="bold"), axis.text.x = element_text(size=10, color="black"), axis.text.y = element_text(size=11, color="black", face="italic"), legend.position="none") + geom_text(aes(label = ..count..), size=5, stat = "count", vjust = 1.5, colour = "white") + scale_x_discrete(labels = function(x) str_wrap(x, width = 7))
dev.off()








