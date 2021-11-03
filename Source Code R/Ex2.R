# Import library
library("readxl")
library(plyr)
library(moments)
library(descr)
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

  ############### Problem No.2 #########################
  
  print(cat("----------------------------------------------------------\nIn", file_name, "\n"))

  clean.data <- subset(data, !is.na(ID) & Status == "Done")   #omit students with n/a IDs and students who haven't done the quiz
  #install.packages("plyr")
  #install.packages("moments")
  #install.packages("descr")
  #CAU A: tinh tong diem tu Q1 den Q10
  K<-clean.data
  print(cat("The Total Score of all Submissions:", sum(K$Total), "\n"))
  
  K.Descend <- K[order(-K$Total),]
  K.Ascend <- K[order(K$Total),]
  #CAU B: Xac dinh diem so thap nhat
  print("   Solving Problem B")
  Least.Total <- min(K[,6]) 
  print(cat("The Least Total Point: ",Least.Total,"\n"))
  
  #CAU C: Xac dinh danh sach cac ban co diem so bang diem so thap nhat
  print("   Solving Problem C")
  List.Least.Total <- subset(K,K$Total == Least.Total )
  List.Least.Total.Unique <- List.Least.Total[match(unique(List.Least.Total$ID),List.Least.Total$ID),]
  print(cat("Student who has at least one total equal to the Least Score: ",List.Least.Total.Unique$ID,"\n"))
  
  #CAU D: Ph??? theo s??? l???n n???p bài c???a các sinh viên có ít nh???t m???t bài có s??? di???m th???p nh???t
  
  print("   Solving Problem D ")
  List.Least.Total2 <-subset(K, ID %in% List.Least.Total.Unique$ID)
  List.Least.Total.Freq <- data.frame(table(List.Least.Total2$ID))
  print(List.Least.Total.Freq)
  barplot(table(List.Least.Total.Freq$Freq),main = "Submissions distribution of students with Least Score", xlab = "Submission", ylab = "Number of students", col = "blue", font = 2)
  #CAU E: XÁC D???NH DI???M S??? T???NG K???T TH???P NH???T 
  print("   Solving Problem E")
  K.Descend.Unique <- K.Descend[match(unique(K.Descend$ID), K.Descend$ID),]
  List.Of.Final.Total <- K.Descend.Unique
  Least.Final.Total <- min(List.Of.Final.Total[,6])
  print(cat("The Least Final Total Score: ",Least.Final.Total,"\n"))

  
  #CAU F:
  print("   Solving Problem F")
  List.Least.Final.Total <- subset(List.Of.Final.Total, Total == Least.Final.Total)
  print(cat("List of Student who has Least Final Total Score: ",List.Least.Final.Total$ID))
  
  #CAU G: Ph??? theo s??? l???n n???p bài c???a các b???n sinh viên có s??? di???m t???ng k???t th???p nh???t
  print("   Solving Problem G ")
  List.Least.Final.Total2 <-subset(K, ID %in% List.Least.Final.Total$ID)
  List.Least.Final.Total.Freq <- data.frame(table(List.Least.Final.Total2$ID))
  barplot(table(List.Least.Final.Total.Freq$Freq),main = "Submissions distribution of students with Least Final Score", xlab = "Submission", ylab = "Number of students", col = "blue", font = 2)
  
  #CAU H:
  print("   Solving Problem H")
  Highest.Total <- max(K[,6]) 
  print(cat("The Highest Total Point: ",Highest.Total,"\n"))
  
  #CAU I:
  print("   Solving Problem I")
  List.Highest.Total <- subset(K,K$Total == Highest.Total )
  List.Highest.Total.Unique <- List.Highest.Total[match(unique(List.Highest.Total$ID),List.Highest.Total$ID),]
  print(cat("Student who has at least one total equal to the Hightest Total Point: ",List.Highest.Total.Unique$ID))
  
  #CAU J:
  print("   Solving Problem J")
  List.Highest.Total2 <-subset(K, ID %in% List.Highest.Total.Unique$ID)
  List.Highest.Total.Freq <- data.frame(table(List.Highest.Total2$ID))
  barplot(table(List.Highest.Total.Freq$Freq),main = "Submissions distribution of students with Highest Score", xlab = "Submission", ylab = "Number of students", col = "blue", font = 2)
  
  #CAU K: XÁC D???NH DI???M S??? T???NG K???T CAO NH???T 
  print("   Solving Problem K")
  Highest.Final.Total <- max(List.Of.Final.Total[,6])
  print(cat("The Highest Final Total Score: ",Highest.Final.Total,"\n"))
  
  #CAU L:
  print("   Solving Problem L")
  List.Highest.Final.Total <- subset(List.Of.Final.Total, Total == Highest.Final.Total)
  print(cat("List of Student who has Highest Final Total Score: ",List.Highest.Final.Total$ID))
  
  #CAU M: Ph??? theo s??? l???n n???p bài c???a các b???n sinh viên có s??? di???m t???ng k???t cao nh???t
  print("   Solving Problem M ")
  List.Highest.Final.Total2 <-subset(K, ID %in% List.Highest.Final.Total$ID)
  List.Highest.Final.Total.Freq <- data.frame(table(List.Highest.Final.Total2$ID))
  barplot(table(List.Highest.Final.Total.Freq$Freq),main = "Submissions distribution of students with Highest Final Score", xlab = "submission", ylab = "Number of students", col = "blue", font = 2)
  
  #CAU N:
  print("   Solving Problem N")
  K.mean <- round(mean(List.Of.Final.Total$Total), digits = 1)
  print(K.mean)
  
  #CAU O:
  print("   Solving Problem O")
  List.mean <-subset(List.Of.Final.Total, Total == K.mean)
  print(cat("The number of student with total score equal to mean:",length(List.mean$ID), "\n"))
  
  #CAU P:
  print(median(K$Total))
  print(min(K$Total))
  print(max(K$Total))
  
  #CAU Q:
  print("   Solving Problem Q")
  print(var(K$Total))
  print(sd(K$Total))
  
  #CAU R:
  print("   Solving Problem R")
  print(skewness(K$Total))
  print(kurtosis(K$Total))
  
  #CAU S:
  print("   Solving Problem S")
  print(quantile(K$Total, 0.25))
  print(quantile(K$Total, 0.75))
  
  #CAU T:
  print("   Solving Problem T")
  Second.Highest.Final.Total <- max(List.Of.Final.Total$Total[List.Of.Final.Total$Total != Highest.Final.Total] )
  List.2Highest.Final.Total <- subset(List.Of.Final.Total, Total == Highest.Final.Total | Total == Second.Highest.Final.Total)
  print(cat("The number of students who has Highest or Second Highest Final Total Score: ",length(List.2Highest.Final.Total$ID), "\n"))
  
  #CAU U:
  print("   Solving Problem U ")
  List.2Highest.Final.Total2 <-subset(K, ID %in% List.2Highest.Final.Total$ID)
  List.2Highest.Final.Total.Freq <- data.frame(table(List.2Highest.Final.Total2$ID))
  barplot(table(List.2Highest.Final.Total.Freq$Freq),main = "Submissions distribution of students with \nHighest or Second Highest Final Score", xlab = "submission", ylab = "Number of students", col = "blue", font = 2)

  #CAU V:
  Total.Descending <- unique(List.Of.Final.Total[order(-List.Of.Final.Total$Total), ]$Total)
  Certain.Total <- Total.Descending[k]
  Certain.List <- subset(List.Of.Final.Total, Total == Certain.Total)
  print(cat("The number of students with Final Total at k-th highest:", nrow(Certain.List), "\n"))

  #CAU W:
  Other.List <- subset(K, ID %in% unique(Certain.List$ID))
  Other.List <- data.frame(table(Other.List$ID))
  barplot(table(Other.List$Freq),main = "Submissions distribution of students with k-Highest Final Score", xlab = "Number of submissions", ylab = "Number of students", col = "blue", font = 2)

}

