# Import library
library("readxl")

# File name list
file_name_list <- c("CO1007_TV_HK192-Quiz 1.4-diem.xlsx",
                    "CO1007_TV_HK192-Quiz 1.5-diem.xlsx",
                    "CO1007_TV_HK192-Quiz 3.3-diem.xlsx",
                    "CO1007_TV_HK192-Quiz 4.2-diem.xlsx")

# Break time to seconds
get_second <- function(string)
{
  time_list = t(as.data.frame(strsplit(string,' ')))
  row.names(time_list) = NULL
  second = 0
  for (i in 1:ncol(time_list)-1)
  {
    if (grepl("ph", time_list[,i+1], fixed = TRUE))
    {
      second = second + as.integer(time_list[,i])*60
    }
    if (grepl("gi", time_list[,i+1], fixed = TRUE))
      second = second + as.integer(time_list[,i])
  }
  result <- second
}

# Break date to int and reconvert
get_date <- function(string)
{
  time_list = t(as.data.frame(strsplit(string,' ')))
  row.names(time_list) = NULL

  month_ = as.integer(time_list[,2])

  hour_minute = t(as.data.frame(strsplit(time_list[,5],':')))
  hour_ = as.integer(hour_minute[,1]) + as.integer(time_list[,6])
  minute_ = as.integer(hour_minute[,2])

  data = data.frame(day = as.integer(time_list[,1]), month = month_, year = as.integer(time_list[,3]), hour = hour_, minute = minute_)
  result <- data
}

get_int_date <- function(list)
{
  int = list[ ,"year"]*10000 + list[ ,"month"]*100 + list[ ,"day"]
  result <- int
}

get_int_date_time <- function(list)
{
  # Base at 0h00 01/03
  day_num = 0
  
  # List for month access
  month_list_day = c(0, 0, 0, 31, 61, 92, 122)

  day_num <- as.integer(month_list_day[as.integer(list[ ,"month"])])

  day_num = day_num + list[ ,"day"] - 1
  int = day_num*1440 + list[ ,"hour"]*60 + list[ ,"minute"]

  result <- int
}

# Read data function
read_data <- function(file_name) 
{
  data <- read_xlsx(file_name, col_names = c("ID", "Status", "Start", "Finish", "Duration",
                                               "Total", "Q1", "Q2", "Q3", "Q4", "Q5", "Q6",
                                               "Q7", "Q8", "Q9", "Q10"))
  # Remove first row
  data <- data[-1,]
  
  # Change format of decimal
  data[] <- lapply(data, function(x) if(grepl(",", x, fixed = TRUE)) as.numeric(sub(",", ".", x)) else x)
  
  # Change format of status
  data[["Status"]] <- lapply(data$Status, function(x) if(grepl(" ho", x, fixed = TRUE)) "Done" else x)
  data[["Status"]] <- lapply(data$Status, function(x) if(grepl("bao", x, fixed = TRUE)) "Not done" else x)
  data["Status"] <- sapply(data$Status, function(x) x[[1]][[1]])

  # Change format of duration
  data[["Duration"]] <- lapply(data$Duration, function(x) as.integer(get_second(x)))
  data["Duration"] <- sapply(data$Duration, function(x) x[[1]][[1]])
  
  # Change format of ID
  data["ID"] <- lapply(data[,"ID"], function(x) as.integer(x))
  
  # Change format of start date
  data[["Start"]] <- lapply(data$Start, function(x) if(grepl("March", x, fixed = TRUE)) sub("March", "3", x) else x)
  data[["Start"]] <- lapply(data$Start, function(x) if(grepl("April", x, fixed = TRUE)) sub("April", "4", x) else x)
  data[["Start"]] <- lapply(data$Start, function(x) if(grepl("May", x, fixed = TRUE)) sub("May", "5", x) else x)
  data[["Start"]] <- lapply(data$Start, function(x) if(grepl("June", x, fixed = TRUE)) sub("June", "6", x) else x)
  data[["Start"]] <- lapply(data$Start, function(x) if(grepl("AM", x, fixed = TRUE)) sub("AM", "0", x) else x)
  data[["Start"]] <- lapply(data$Start, function(x) if(grepl("PM", x, fixed = TRUE)) sub("PM", "12", x) else x)
  data[["Start"]] <- lapply(data$Start, function(x) if(grepl("12:", x, fixed = TRUE)) sub("12:", "0:", x) else x)
  data["Start"] <- sapply(data$Start, function(x) x[[1]][[1]])
  data["Start"] <- lapply(data[,"Start"], function(x) get_date(x))
  
  # Change format of Finish date
  data[["Finish"]] <- lapply(data$Finish, function(x) if(grepl("March", x, fixed = TRUE)) sub("March", "3", x) else x)
  data[["Finish"]] <- lapply(data$Finish, function(x) if(grepl("April", x, fixed = TRUE)) sub("April", "4", x) else x)
  data[["Finish"]] <- lapply(data$Finish, function(x) if(grepl("May", x, fixed = TRUE)) sub("May", "5", x) else x)
  data[["Finish"]] <- lapply(data$Finish, function(x) if(grepl("June", x, fixed = TRUE)) sub("June", "6", x) else x)
  data[["Finish"]] <- lapply(data$Finish, function(x) if(grepl("AM", x, fixed = TRUE)) sub("AM", "0", x) else x)
  data[["Finish"]] <- lapply(data$Finish, function(x) if(grepl("PM", x, fixed = TRUE)) sub("PM", "12", x) else x)
  data[["Finish"]] <- lapply(data$Finish, function(x) if(grepl("12:", x, fixed = TRUE)) sub("12:", "0:", x) else x)
  data["Finish"] <- sapply(data$Finish, function(x) x[[1]][[1]])
  data["Finish"] <- lapply(data[,"Finish"], function(x) get_date(x))
  
  result <- data
}

# Ignore warning
options(warn = -1)

################################################################################################################
############ Main program ######################################################################################

# Insert any additional functions here

for (file_name in file_name_list) 
{
  data <- read_data(file_name)
  
  ####### Problem no.12 - Duration analysis ############
  data <- subset(data, Status == "Done")
  
  print(cat("------------------------------------------------------------------\nIn", file_name, "\n"))
  
  
  ################################## Normal ###############################################################
  # Histogram of duration
  hist(data$Duration, main = "Histogram of duration", xlab = "Duration (sec)", ylab = "Number of submissions",
       col = "green", breaks = ((0:150)/150)*max(data$Duration))
  
  # Mean of duration
  print(cat("Average of duration (sec):", mean(data$Duration), "\n"))
  
  ################################## First Submission ####################################################
  # Histogram of duration of the first submission
  first_submit <- data[order(data$Start$year, data$Start$month, data$Start$day, data$Start$hour, data$Start$minute),]
  first_submit <- first_submit[match((unique(first_submit$ID)), first_submit$ID),]
  hist(first_submit$Duration, main = "Histogram of duration of first submissions", xlab = "Duration (sec)", ylab = "Number of submissions",
       col = "green", breaks = ((0:150)/150)*max(data$Duration))
  
  # Mean of duration and total score of first submissions
  print(cat("Average of duration of first submissions (sec):", mean(first_submit$Duration), "\n"))
  print(cat("Average of total score of first submissions (sec):", mean(first_submit$Total), "\n"))
  
  ################################## Last Submission ####################################################
  # Histogram of duration of the last submissions
  last_submit <- data[order(-data$Start$year, -data$Start$month, -data$Start$day, -data$Start$hour, -data$Start$minute),]
  last_submit <- last_submit[match((unique(last_submit$ID)), last_submit$ID),]
  hist(last_submit$Duration, main = "Histogram of duration of last submissions", xlab = "Duration (sec)", ylab = "Number of submissions",
       col = "green", breaks = ((0:150)/150)*max(data$Duration))
  
  # Mean of duration and total score of last submissions
  print(cat("Average of duration of last submissions (sec):", mean(last_submit$Duration), "\n"))
  print(cat("Average of total score of last submissions (sec):", mean(last_submit$Total), "\n"))
  
  ################################## Last Submission Fake ####################################################
  # Histogram of duration of the last submissions of students with 2 or more submissions
  last_fake <- data[order(data$Start$year, data$Start$month, data$Start$day, data$Start$hour, data$Start$minute),]
  last_fake <- last_fake[-match((unique(last_fake$ID)), last_fake$ID),]
  
  last_fake <- last_fake[order(-last_fake$Start$year, -last_fake$Start$month, -last_fake$Start$day, -last_fake$Start$hour, -last_fake$Start$minute),]
  last_fake <- last_fake[match((unique(last_fake$ID)), last_fake$ID),]
  hist(last_fake$Duration, main = "Histogram of duration of last submissions \nof students with 2 or more submissions", 
       xlab = "Duration (sec)", ylab = "Number of submissions",
       col = "green", breaks = ((0:150)/150)*max(data$Duration))
  
  # Mean of duration and total score of last submissions of students with 2 or more submissions
  print(cat("Average of duration of last submissions of students with 2 or more submissions:", mean(last_fake$Duration), "\n"))
  print(cat("Average of total score of last submissions of students with 2 or more submissions:", mean(last_fake$Total), "\n"))
  
  # Histogram of score in this case
  hist(last_fake$Total, main = "Histogram of score of last submissions \nof students with 2 or more submissions", 
       xlab = "Total score", ylab = "Number of submissions",
       col = "blue", breaks = (0:20)*0.5)
  
  ################################## First Submission Fake ##################################################
  first_fake <- subset(first_submit, ID %in% last_fake$ID)
  
  # Histogram of duration
  hist(first_fake$Duration, main = "Histogram of duration of first submissions \nof students with 2 or more submissions", 
       xlab = "Duration (sec)", ylab = "Number of submissions",
       col = "green", breaks = ((0:150)/150)*max(data$Duration))

  # Mean of duration and total score of last submissions of students with 2 or more submissions
  print(cat("Average of duration of first submissions of students with 2 or more submissions:", mean(first_fake$Duration), "\n"))
  print(cat("Average of total score of first submissions of students with 2 or more submissions:", mean(first_fake$Total), "\n"))

  # Histogram of score
  hist(first_fake$Total, main = "Histogram of score of first submissions \nof students with 2 or more submissions", 
       xlab = "Total score", ylab = "Number of submissions",
       col = "blue", breaks = (0:20)*0.5)
  
  ################################# Difference in duration and score ########################################
  difference <- merge(first_fake, last_fake, by = "ID")[, c("ID", "Duration.x", "Duration.y", "Total.x", "Total.y")]
  difference$DurDif <- (difference$Duration.x - difference$Duration.y)
  difference$TotalDif <- (difference$Total.y - difference$Total.x)
  
  print(cat("Average difference in duration between first and last submissions of students with 2 or more submissions (sec):",
            mean(difference$DurDif), "\n"))
  print(cat("Average difference in total score between first and last submissions of students with 2 or more submissions:",
            mean(difference$TotalDif), "\n"))
}

