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
  
  ###### Problem no. 10 #######
  
  print("-------------------------------------------------------------------------------------------")
  print(file_name)
  
  # Index a) Determine suitable arguments
  actual.data <- subset(data, Status == "Done")

  ### Determine smart students
  num.of.sub <- data.frame(table(actual.data$ID))
  avg.num.of.sub <- floor(mean(num.of.sub$Freq))
  
  the.elite <- actual.data[order(actual.data$ID, -actual.data$Total), ]
  the.elite <- the.elite[match(unique(the.elite$ID), the.elite$ID), ]
  the.elite <- the.elite[order(-the.elite$Total), ]
  required.total <- min(the.elite$Total[1 : (length(the.elite$Total) / 10)])

  temp.data <- actual.data[order(actual.data$Start$year, actual.data$Start$month, actual.data$Start$day, actual.data$Start$hour, actual.data$Start$minute), ]
  final.data <- actual.data[-(1:nrow(actual.data)), ]

  for (i in 1:avg.num.of.sub)
  {
  	match1 <- temp.data[match(unique(temp.data$ID), temp.data$ID), ]
  	temp.data <- temp.data[-match(unique(temp.data$ID), temp.data$ID), ]
  	final.data <- rbind(final.data, match1)
  }

  final.data <- final.data[order(-final.data$Total), ]
  final.data <- final.data[match(unique(final.data$ID), final.data$ID), ]
  final.data <- subset(final.data, Total >= required.total)
  
  ### Determine hard-working students
  the.hard <- actual.data[order(actual.data$Finish$year, actual.data$Finish$month, actual.data$Finish$day, actual.data$Finish$hour, actual.data$Finish$minute), ]
  the.hard <- the.hard[(1:(nrow(the.hard) %/% 10)), ]
  
  required.time <- the.hard[nrow(the.hard), "Finish"]
  
  the.hard <- the.hard[order(-the.hard$Total), ]
  the.hard <- the.hard[match(unique(the.hard$ID), the.hard$ID), ]

  req.num.of.sub <- ceiling(mean(num.of.sub$Freq))
  req.hard <- subset(num.of.sub, Freq >= req.num.of.sub)

  the.hard.active <- subset(the.hard, ID %in% unique(req.hard$Var1))

  ### Determine the union
  final.data <- rbind(final.data, the.hard.active)
  final.data <- final.data[order(-final.data$Total), ]
  final.data <- final.data[match(unique(final.data$ID), final.data$ID), ]
  
  ### Print arguments
  print(cat("Smart students argument:", "maximal number of submissions:", avg.num.of.sub, "; required total score:",
            required.total, "\n", "Hard-working students argument:", "minimal number of submissions:", req.num.of.sub,
            "; latest Finish time:", required.time$Finish$hour[1], ":", required.time$Finish$minute[1], 
            ",", "day:", required.time$Finish$day[1], ",", "month:", required.time$Finish$month[1], ",", "year:", 
            required.time$Finish$year[1], "\n"))
  
  # Index b) Determine the number of active students
  print(cat("The number of active students:", nrow(final.data), "\n"))
  
  # Index c) Draw distribution of total score of active students
  hist(final.data$Total, main = "Histogram of active students' total score", xlab = "Score", ylab = "Number of students",
       col = "orange", xlim = c(0,10), breaks = ((0:20)*0.5))
  
}
