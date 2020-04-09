# clear
rm(list = ls())

# packages
require(tidyverse)
require(xlsx)
require(openxlsx)

# data
df <- read.csv("C:/Users/moctar.aboubacar/World Food Programme/SO5- Evidence Policy & Innovation - COVID-19 response/mVAM/Study design/Phone numbers first digits.csv")

df$number <- as.character(df$number)
df$zone_num <- as.factor(df$zone_num)
glimpse(df)

# create possible back digits
vec <- as.character(1:100000)
vec <- str_pad(vec,
               side = "left",
               width = 5,
               pad = "0")


# create all numbers
all_num <- function(df, vec){
  
mat <- matrix(ncol = length(df$number),
              nrow = length(vec))

for (i in 1:nrow(df)) {
  temp <- paste0(df$number[i], vec)
  mat[,i] <- temp
}

return(mat)
}

numbers <- all_num(df, vec)


# sample according to design
key <- read.csv("C:/Users/moctar.aboubacar/World Food Programme/SO5- Evidence Policy & Innovation - COVID-19 response/mVAM/Study design/sample size by type.csv")

zone_sims <- function(zone = "Bagmati"){

matree <- matrix(ncol = 10,
                 nrow = (as.integer(key[key$zone == zone & key$type == "NT" & key$prepost == "pre",][5]) + as.integer(key[key$zone == zone & key$type == "NT" & key$prepost == "post",][5]) + as.integer(key[key$zone == zone & key$type == "Ncell" & key$prepost == "pre",][5]) + as.integer(key[key$zone == zone & key$type == "Ncell" & key$prepost == "post",][5])))

for (i in 1:10) {
v1 <- sample(numbers[,which(df$zone == zone & df$type == "NT" & df$prepost == "pre")],
             size = as.integer(key[key$zone == zone & key$type == "NT" & key$prepost == "pre",][5]))

v2 <- sample(numbers[,which(df$zone == zone & df$type == "NT" & df$prepost == "post")],
             size = as.integer(key[key$zone == zone & key$type == "NT" & key$prepost == "post",][5]))

v3 <- sample(numbers[,which(df$zone == zone & df$type == "Ncell" & df$prepost == "pre")],
             size = as.integer(key[key$zone == zone & key$type == "Ncell" & key$prepost == "pre",][5]))

v4 <- sample(numbers[,which(df$zone == zone & df$type == "Ncell" & df$prepost == "post")],
             size = as.integer(key[key$zone == zone & key$type == "Ncell" & key$prepost == "post",][5]))

matree[,i] <- sample(c(v1, v2, v3, v4))
}
matree
}

# construct matrices of results
matList <- map(levels(df$zone), zone_sims)
names(matList) <- levels(df$zone)

# save to excel file in 3 steps

# 1 create workbook
numbers <- createWorkbook()

# 2 iterate with anonymous function inside Map())
Map(function(matList, nameofsheet){     
  
  addWorksheet(numbers, nameofsheet)
  writeData(numbers, nameofsheet, matList)
  
}, matList, names(matList))

# 3 save workbook to excel file 
saveWorkbook(numbers, file = "C:/Users/moctar.aboubacar/World Food Programme/SO5- Evidence Policy & Innovation - COVID-19 response/mVAM/Study design/numbers.xlsx", overwrite = TRUE)