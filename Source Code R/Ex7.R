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
  
  ###### Problem no. 7 #######
  
  print(file_name)
  
  # Index a) Determine t_2
  actual_data <- subset(data, Status == "Done")
  
  the_lazy_as <- actual_data[order(actual_data$Start$year, actual_data$Start$month, actual_data$Start$day, actual_data$Start$hour, actual_data$Start$minute), ]
  
  ## Remove second submission onwards
  the_lazy_as <- the_lazy_as[match(unique(the_lazy_as$ID), the_lazy_as$ID), ]
  
  ## Get 10% students who has latest first submission
  the_lazy_as <- the_lazy_as[(nrow(the_lazy_as) - (nrow(the_lazy_as) %/% 10)):nrow(the_lazy_as), ]
  
  ## Get time_2
  time_2 <- the_lazy_as[1,"Start"]
  
  ## Print time_2
  print(cat("The suitable time:", time_2$Start$hour[1], ":", time_2$Start$minute[1], ",", "day:",
            time_2$Start$day[1], ",", "month:", time_2$Start$month[1], ",", "year:", time_2$Start$year[1], "\n"))
  
  # Index b) Determine the number of lazy students
  print(cat("The number of lazy students:", nrow(the_lazy_as), "\n"))
  
  # Index c) Draw histogram of total score of lazy students
  hist(the_lazy_as$Total, main = "Histogram of lazy students' scores (first submission)",
       xlab = "Total score", ylab = "Number of students", xlim = c(0,10),
       breaks = ((0:20)*0.5))
  
  the_lazy_des <- subset(actual_data, ID %in% unique(the_lazy_as$ID))
  
  the_lazy_des <- the_lazy_des[order(-the_lazy_des$Total), ]
  
  ## Get highest total score
  the_lazy_des <- the_lazy_des[match(unique(the_lazy_des$ID), the_lazy_des$ID), ]
  
  hist(the_lazy_des$Total, main = "Histogram of lazy students' scores (highest score)",
       xlab = "Total score", ylab = "Number of students", xlim = c(0,10), breaks = ((0:20)*0.5))
}
