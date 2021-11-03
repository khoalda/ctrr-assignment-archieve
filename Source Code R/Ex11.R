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
  
  actual.data <- subset(data, Status == "Done")
  actual.stu <- actual.data[match(unique(actual.data$ID), actual.data$ID), ]

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

  smart.list <- unique(final.data$ID)
  
  ### Determine hard-working students
  the.hard <- actual.data[order(actual.data$Start$year, actual.data$Start$month, actual.data$Start$day, actual.data$Start$hour, actual.data$Start$minute), ]
  the.hard <- the.hard[(1:(nrow(the.hard) %/% 10)), ]
  
  required.time <- the.hard[nrow(the.hard), "Start"]
  
  hard.list <- unique(the.hard$ID)
  
  the.hard <- the.hard[order(-the.hard$Total), ]
  the.hard <- the.hard[match(unique(the.hard$ID), the.hard$ID), ]

  req.num.of.sub <- ceiling(mean(num.of.sub$Freq))
  req.hard <- subset(num.of.sub, Freq >= req.num.of.sub)

  the.hard.active <- subset(the.hard, ID %in% unique(req.hard$Var1))

  ### Determine the union - active students
  final.data <- rbind(final.data, the.hard.active)
  final.data <- final.data[order(-final.data$Total), ]
  final.data <- final.data[match(unique(final.data$ID), final.data$ID), ]

  active.list <- unique(final.data$ID)

  ### Determine the lazy
  the.lazy.as <- actual.data[order(actual.data$Start$year, actual.data$Start$month, actual.data$Start$day, actual.data$Start$hour, actual.data$Start$minute), ]
  the.lazy.as <- the.lazy.as[match(unique(the.lazy.as$ID), the.lazy.as$ID), ]
  the.lazy.as <- the.lazy.as[(nrow(the.lazy.as) - (nrow(the.lazy.as) %/% 10)):nrow(the.lazy.as), ]

  lazy.list <- unique(the.lazy.as$ID)

  ### Determine the good 
  the.good <- actual.data[order(-actual.data$Total), ]
  the.good <- the.good[match(unique(the.good$ID), the.good$ID), ]
  required.total <- the.good$Total[nrow(the.good) %/% 5]

  the.good <- subset(the.good, Total >= required.total)

  good.list <- unique(the.good$ID)
  
  ### Active: AC, Smart: SM, Good: GO, Lazy: LA, Hard-working: HA
  print(cat("The number of students in each group: AC - SM - GO - LA - HA", "\n",
            "Active (AC):", nrow(subset(actual.stu, ID %in% active.list)), "\n",
            "Smart (SM):", nrow(subset(actual.stu, ID %in% smart.list)), "\n",
            "Good (GO):", nrow(subset(actual.stu, ID %in% good.list)), "\n",
            "Lazy (LA):", nrow(subset(actual.stu, ID %in% lazy.list)), "\n",
            "Hard-working (HA):", nrow(subset(actual.stu, ID %in% hard.list)), "\n",
            
            "AC ^ SM:", nrow(subset(actual.stu, (ID %in% active.list) & (ID %in% smart.list))), "\n",
            "AC ^ GO:", nrow(subset(actual.stu, (ID %in% active.list) & (ID %in% good.list))), "\n",
            "AC ^ LA:", nrow(subset(actual.stu, (ID %in% active.list) & (ID %in% lazy.list))), "\n",
            "AC ^ HA:", nrow(subset(actual.stu, (ID %in% active.list) & (ID %in% hard.list))), "\n",
            "SM ^ GO:", nrow(subset(actual.stu, (ID %in% smart.list) & (ID %in% good.list))), "\n",
            "SM ^ LA:", nrow(subset(actual.stu, (ID %in% smart.list) & (ID %in% lazy.list))), "\n",
            "SM ^ HA:", nrow(subset(actual.stu, (ID %in% smart.list) & (ID %in% hard.list))), "\n",
            "GO ^ LA:", nrow(subset(actual.stu, (ID %in% good.list) & (ID %in% lazy.list))), "\n",
            "GO ^ HA:", nrow(subset(actual.stu, (ID %in% good.list) & (ID %in% hard.list))), "\n",
            "LA ^ HA:", nrow(subset(actual.stu, (ID %in% lazy.list) & (ID %in% hard.list))), "\n",
            
            "SM ^ GO ^ LA:", nrow(subset(actual.stu, (ID %in% smart.list)  & (ID %in% good.list) & (ID %in% lazy.list))), "\n",
            "SM ^ GO ^ HA:", nrow(subset(actual.stu, (ID %in% smart.list)  & (ID %in% good.list) & (ID %in% hard.list))), "\n",
            "AC ^ GO ^ LA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% good.list) & (ID %in% lazy.list))), "\n",
            "SM ^ LA ^ HA:", nrow(subset(actual.stu, (ID %in% smart.list)  & (ID %in% lazy.list) & (ID %in% hard.list))), "\n",
            "AC ^ GO ^ HA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% good.list) & (ID %in% hard.list))), "\n",
            "AC ^ SM ^ LA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% smart.list) & (ID %in% lazy.list))), "\n",
            "GO ^ LA ^ HA:", nrow(subset(actual.stu, (ID %in% good.list)  & (ID %in% lazy.list) & (ID %in% hard.list))), "\n",
            "AC ^ LA ^ HA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% lazy.list) & (ID %in% hard.list))), "\n",
            "AC ^ SM ^ HA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% smart.list) & (ID %in% hard.list))), "\n",
            "AC ^ SM ^ GO:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% smart.list) & (ID %in% good.list))), "\n",
            
            "AC ^ SM ^ LA ^ HA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% smart.list) & (ID %in% lazy.list) & (ID %in% hard.list))), "\n",
            "AC ^ GO ^ LA ^ HA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% good.list) & (ID %in% lazy.list) & (ID %in% hard.list))), "\n",
            "AC ^ SM ^ GO ^ HA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% smart.list) & (ID %in% good.list) & (ID %in% hard.list))), "\n",
            "AC ^ SM ^ GO ^ LA:", nrow(subset(actual.stu, (ID %in% active.list)  & (ID %in% smart.list) & (ID %in% good.list) & (ID %in% lazy.list))), "\n",
            "SM ^ GO ^ LA ^ HA:", nrow(subset(actual.stu, (ID %in% smart.list)  & (ID %in% good.list) & (ID %in% lazy.list) & (ID %in% hard.list))), "\n",
            
            "Inteersection of all:", nrow(subset(actual.stu, (ID %in% active.list) & (ID %in% smart.list) & (ID %in% good.list) & (ID %in% lazy.list) & (ID %in% hard.list))), "\n"
            ))
  
}
