# Import library
library("readxl")
library(plyr)
library(moments)

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

k <- readline(prompt="Enter k: ")
k <- as.integer(k)
if (is.na(k))
{
	k <- 2
	print("Set k = 2\n")
}

for (file_name in file_name_list) 
{
  data <- read_data(file_name)
  
  ########## Problem no.3 ############
  
  print(cat("In", file_name, "\n"))
  
  clean_data <- subset(data, !is.na(ID) & Status == "Done")   #omit students with n/a IDs and students who haven't done the quiz
  #install.packages("plyr")
  #install.packages("moments")
  
  
  submission_table <- count(clean_data$ID) #count frequency of each ID -> number of submissions
  
  min_num <- min(submission_table$freq)
  max_num <- max(submission_table$freq)
  avg_num <- round(mean(submission_table$freq))
  second_max_num <- max(submission_table$freq[submission_table$freq != max_num])
  
  filtered_data <- clean_data[!rev(duplicated(rev(clean_data$ID))),]
  arranged_data <- filtered_data[order(filtered_data$ID),]
  arranged_data$submission = submission_table$freq
  
  least_subset <- subset(arranged_data, submission == min_num)
  most_subset <- subset(arranged_data, submission == max_num)
  avg_subset <- subset(arranged_data, submission == avg_num)
  most_2nd_subset <- subset(arranged_data, submission == second_max_num)
  most_group_subset <- subset(arranged_data, submission == max_num  | submission == second_max_num)
  
  print(cat("a. The smallest number of submissions:", min_num, "\n"))
  print(cat("b. The students with least submissions:", least_subset$ID, "\n"))
  print(cat("c. Score distribution of the students with least submissions:", "\n"))
  table(least_subset$Total)
  barplot(table(least_subset$Total), main = "Score distribution of the students with \nleast submissions", xlab = "Score", ylab = "Count", col = "red", font = 2)
  
  print(cat("d. The largest number of submissions:", max_num, "\n"))
  print(cat("e. The students with most submissions:", most_subset$ID, "\n"))
  print(cat("f. Score distribution of the students with most submissions:", "\n"))
  table(most_subset$Total)
  barplot(table(most_subset$Total), main = "Score distribution of the students with \nmost submissions", xlab = "Score", ylab = "Count", col = "red", font = 2)
  
  print(cat("g. The average number of submissions:", avg_num, "\n"))
  print(cat("h. The number of students with average submissions:", length(avg_subset$ID), "\n"))
  print(cat("i. Score distribution of the students with average submissions:", "\n"))
  table(avg_subset$Total)
  barplot(table(avg_subset$Total), main = "Score distribution of the students with \naverage submissions", xlab = "Score", ylab = "Count", col = "red", font = 2)
  print(cat("j. The median of sample above:", median(avg_subset$Total), "\nThe maxima of sample above:", max(avg_subset$Total),"\nThe minima of sample above:", min(avg_subset$Total),"\n"))
  
  print(cat("k. The variance of sample above:", var(avg_subset$Total), "\n", "   The standard deviation of sample above:", sd(avg_subset$Total), "\n"))
  print(cat("l. The skewness of data in sample above:", skewness(avg_subset$Total), "\nThe kurtosis of data in sample above:", kurtosis(avg_subset$Total),"\n"))
  print(cat("m. Q1 of sample above:", quantile(avg_subset$Total, 0.25, na.rm = TRUE), "\nQ3 of sample above:", quantile(avg_subset$Total, 0.75, na.rm = TRUE),"\n"))
  
  print(cat("n. The students with second most submissions:", most_2nd_subset$ID, "\n"))
  
  print(cat("o. The students with most or second most submissions:", most_group_subset$ID, "\n"))
  print(cat("p. The number of students with most or second most submissions:", length(most_group_subset$ID), "\n"))
  print(cat("q. Score distribution of the students with most or second most submissions:", "\n"))
  table(most_group_subset$Total)
  barplot(table(most_group_subset$Total), main = "Score distribution of the students with \nmost or second most submissions", xlab = "Score", ylab = "Count", col = "red", font = 2)
  
  decreasing_data <- arranged_data[order(arranged_data$submission, decreasing = TRUE),]
  top_third <- head(decreasing_data, length(decreasing_data$submission)/3)
  print(cat("r. The students in the top third:", top_third$ID, "\n"))
  print(cat("s. The number of students in the top third:", length(top_third$ID), "\n"))
  print(cat("t. Score distribution of the students in the top third:", "\n"))
  table(top_third$Total)
  barplot(table(top_third$Total), main = "Score distribution of the students in the top third", xlab = "Score", ylab = "Count", col = "red", font = 2)
  
  #k is input from keyboard
  max_temp <- max_num
  if (k == 1){
    k_th_max <- max_num
  }else{
    for (i in 2:k) 
    {
      k_th_max <- max(submission_table$freq[submission_table$freq < max_temp])
      max_temp <- k_th_max
    }
  }
  
  top_k_group <- subset(arranged_data, submission >= k_th_max)
  print(cat("u. Score distribution of the students in the top", k, "based on number of submissions: "))
  table(top_k_group$Total)
  barplot(table(top_k_group$Total), main = "Score distribution of the students in the top k \nbased on number of submissions", xlab = "Score", ylab = "Count", col = "red", font = 2)
  
  print(cat("_______________________________________________", "\n"))
}
