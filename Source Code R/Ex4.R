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
  int = list[1,"year"]*10000 + list[1,"month"]*100 + list[1,"day"]
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
  data <- subset(data, Status == "Done")
  
  
  print(cat("-----------------------------------------------------------------\nIn", file_name, "\n"))
  
  
  #CAU a
  nopbailandau<-data.frame(data[,1],data[,4])
  nopbailandau<-nopbailandau[order(nopbailandau$Finish$year, nopbailandau$Finish$month, nopbailandau$Finish$day, nopbailandau$Finish$hour, nopbailandau$Finish$minute),]
  nopbailandau<-nopbailandau[match(unique(nopbailandau$ID),nopbailandau$ID),]
  nopbailandau<-nopbailandau[order(nopbailandau$ID),]
  nopbailancuoi<-data.frame(data[,1],data[,4])
  nopbailancuoi<-nopbailancuoi[order(-nopbailancuoi$Finish$year, -nopbailancuoi$Finish$month, -nopbailancuoi$Finish$day, -nopbailancuoi$Finish$hour, -nopbailancuoi$Finish$minute),]
  nopbailancuoi<-nopbailancuoi[match(unique(nopbailancuoi$ID),nopbailancuoi$ID),]
  nopbailancuoi<-nopbailancuoi[order(nopbailancuoi$ID),]
  tgnopbai<-data.frame(nopbailancuoi$ID)
  tgnopbai[["land"]]<-nopbailandau$Finish

  tgnopbai[["lanc"]]<-nopbailancuoi$Finish
  tgnopbai$tg <- get_int_date_time(tgnopbai$lanc) - get_int_date_time(tgnopbai$land)
  
  ketquacau1<-data.frame(tgnopbai$nopbailancuoi.ID,tgnopbai$tg)
  print(cat("Maximal duration between last and first submissions:", max(ketquacau1$tgnopbai.tg), "\n"))
  
  
  
  #CAU b
  hist(tgnopbai$tg, main="Histogram of students' working time", xlab="Time (min)",ylab = "Number of students")
  
  
  
  #CAU c

  newketqua<-subset(ketquacau1,tgnopbai.tg >=0)
  ts <- data.frame(table(data$ID))
  #ts$time <- newketqua$tgnopbai.tg
  ts$tansuatnopbai<-(newketqua$tgnopbai.tg) / (ts$Freq)
  ketquacau3<-data.frame(ts$Var1,ts$tansuatnopbai)
  print("List of students with their duration per submission:")
  print(ketquacau3)


  #CAU d
  ketquacau4<-subset(ketquacau3,ts.tansuatnopbai==min(ketquacau3$ts.tansuatnopbai))
  print("List of students with lowest duration per submission:")
  print(ketquacau4$ts.Var1)


  #CAU e: xac dinh pho diem cua sinh vien co tan suat nop bai it nhat
  rut<-data.frame(data$ID,data$Total)
  rut<-rut[order(-rut$data.Total),]
  rut<-rut[match(unique(rut$data.ID),rut$data.ID),]
  ketquacau5 <- subset(rut, data.ID %in% ketquacau4$ts.Var1)
  hist(ketquacau5$data.Total,main="Histogram of scores of students with lowest \nsubmitting relative frequency", xlab="Score",ylab = "Number of students", breaks = ((0:20)*0.5))


  #CAU f: xac dinh so luong sinh vien co tan suat nop bai nhieu nhat
  ketquacau6<-subset(ketquacau3,ts.tansuatnopbai==max(ketquacau3$ts.tansuatnopbai))
  print(cat("The number of students with highest duration per permission:",length(ketquacau6$ts.Var1),"\n"))

  #CAU g: xac dinh sinh vien co tan suat nop bai nhieu nhat
  ketquacau7<-subset(ketquacau3,ts.tansuatnopbai==max(ketquacau3$ts.tansuatnopbai))
  print("List of students with highest duration per submission:")
  print(ketquacau7$ts.Var1)


  #CAU h: xac dinh pho diem cua sinh vien co tan suat nop bai nhieu nhat
  rut<-data.frame(data$ID,data$Total)
  rut<-rut[order(-rut$data.Total),]
  rut<-rut[match(unique(rut$data.ID),rut$data.ID),]
  ketquacau8 <- subset(rut, data.ID %in% ketquacau7$ts.Var1)
  hist(ketquacau8$data.Total,main="Histogram of scores of students with highest \nsubmitting relative frequency", xlab="Score",ylab = "Number of students", breaks = ((0:20)*0.5))
  
  
  #CAU i: xac dinh cac sinh vien co tan suat nop bai nhieu nhi
  ketquacau9<-subset(ketquacau3,ts.tansuatnopbai<max(ketquacau3$ts.tansuatnopbai))
  ketquacau9<-subset(ketquacau9,ts.tansuatnopbai==max(ketquacau9$ts.tansuatnopbai))
  print("List of students with 2nd-highest duration per submission:")
  print(ketquacau9$ts.Var1)
  
  
  
  #CAU j: xac dinh sinh vien nam trong nhom co tan suat nop bai nhieu nhat va nhi
  ketquacau10<-subset(ketquacau3,ts.tansuatnopbai>=max(ketquacau9$ts.tansuatnopbai))
  
  print("List of students with highest or 2nd-highest duration per submission:")
  print(ketquacau10$ts.Var1)
  
  #CAU k: xac dinh thoi gian giua 2 lan nop lien tiep
  ts$tansuatnopbaigiay <-(newketqua$tgnopbai.tg*60) / (ts$Freq-1)
  ketquacau11<-data.frame(ts$Var1,ts$tansuatnopbaigiay)
  ketquacau11<-subset(ketquacau11, ts.tansuatnopbaigiay>=0)
  print("List of students with their average duration between 2 consecutive submissions:")
  print(ketquacau11)
 
  
   #CAU l: tinh tan so, tan suat va tan suat tich luy
  tanso<-data.frame(table(ketquacau3$ts.tansuatnopbai))
  print("Duration per submission frequency tables:")
  print(tanso)

  tansuat<- table((ketquacau3$ts.tansuatnopbai))/length(ketquacau3$ts.tansuatnopbai)

  print("Duration per submission relative frequency tables:")
  print(data.frame(tansuat))

  x=table(cut(ketquacau3$ts.tansuatnopbai,breaks=c(0:10)))
  tansuatluytich <- cumsum(x)/length(ketquacau3$ts.tansuatnopbai)
  print("Duration per submission cumulative frequency tables:")
  print(data.frame(tansuatluytich))
  
  
  
  
  #CAU m: ve bieu do tan so
  hist(ketquacau3$ts.tansuatnopbai,main="Frequency of duration per submission", xlab="Duration per submission",ylab = "Frequency")
  
  
  
  
  #CAU n: ve bieu do tan suat
  barplot(tansuat, beside=TRUE,main="Relative frequency of duration per submission", xlab="Duration per submission", ylab="Relative frequency")
  
  
  
  
  #CAU o: ve bieu do tan suat tich luy
  barplot(tansuatluytich, beside=TRUE,main="Cumulative frequency of duration per submission", xlab="Duration per submission", ylab="Cumulative frequency")
  
  
  
  
  
  #CAU p: tinh trung vi mau, cuc dai mau, cuc tieu mau
 
  
  print(cat("Median of sample:", median(ketquacau3$ts.tansuatnopbai),"\n"))
  print(cat("Maximal of sample:", max(ketquacau3$ts.tansuatnopbai),"\n"))
  print(cat("Minimal of sample:", min(ketquacau3$ts.tansuatnopbai),"\n"))
 
  
  
  
  
  #CAU q: tim do phan tan cua tan suat nop bai (ta dua vao phuong sai va do lech chuan)
  
  print(cat("Variance of sample:",var(ketquacau3$ts.tansuatnopbai), "- standard deviation", sd(ketquacau3$ts.tansuatnopbai), "\n"))
  
  
  
  
  #CAU r: Tinh do lech, va do nhon
  print(cat("Skewness of sample:",skewness(ketquacau3$ts.tansuatnopbai),"\n"))
  print(cat("Kurtosis of sample:",kurtosis(ketquacau3$ts.tansuatnopbai),"\n"))
  
  
  
  
  
  #CAU s: tinh tu phan vi thu nhat va thu 3
  
  print(cat("First quantile of the sample:",quantile(ketquacau3$ts.tansuatnopbai,0.25),"\n"))
  print(cat("Third quantile of the sample:",quantile(ketquacau3$ts.tansuatnopbai,0.75),"\n"))
}

