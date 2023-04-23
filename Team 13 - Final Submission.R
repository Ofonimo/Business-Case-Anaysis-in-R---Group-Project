################################
### Title  : A1 - Business Case Presentation
### Team   : 13
### Members: Alberto Valencia, Htet Aung Kyaw, Ofonimo Ben, 
###          Ryosuke Ogata, Sanskriti Garg
### Course : Visualizing & Analyzing Data with R: Methods & Tools
### Prof.  : Thomas Kurnicki
################################

# Install requires libraries
# install.packages("readxl")
# install.packages("tidyverse")
# install.packages("openxlsx")

# Load required libraries
library(readxl)
library(tidyverse)
library(openxlsx)

# Declare file path
file_path <- "./Air France Case Spreadsheet Supplement.xls"

# Function to read excel data
read_excel_data <- function(file_path) {
  df <- read_xls(path = file_path, col_names = TRUE)
  return(df)
}

# Function to check missing values
check_missing_values <- function(df) {
  for (i in 1:ncol(df)){
    cat(sum(is.na(df[,i]) == TRUE), "-", colnames(df[i]), "\n")
  }
}
# the Bid Strategy column contains missing values

# Function to normalize variables
normalize <- function(x, min, max) {
  return((x - min) / (max - min))
}


################## Exploration and Visualization begin ####################

### Explore, analyzing publisher and campaign before optimization
# group by publisher name
# read air France excel file and store it to a object called af_df
# sheet 2 with column names
# use read excel library
library(readxl)

# read air France excel file and store it to a object called af_df
# sheet 2 with column names
af_df <- read_xls(path="./Air France Case Spreadsheet Supplement.xls",
                  sheet=2,
                  col_names = TRUE)
# check summary of air France data frame
summary(af_df)

normalize <- function(x, min, max){ # x is the variable vector
  min_max <- ((x- min)/(max- min))
  return(min_max)
}# closing normalize function


# check missing value using for loop
for (i in 1:ncol(af_df)){
  cat(sum(is.na(af_df[,i]) == TRUE), "-", colnames(af_df[i]), "\n")
}

# user tidyverse package for dylyr and ggplot 
library(tidyverse)

# group by publisher name
group_by_pub <- af_df %>% 
  group_by(`Publisher Name`) %>% 
  summarise(cost_per_click = sum(`Click Charges`) / sum(`Clicks`),
            click_thru_rate = round((sum(`Clicks`)/ sum(Impressions) ) * 100, 2),
            total_booking_per_click = round((sum(`Total Volume of Bookings`) / sum(Clicks))*100,2),
            cus_acq_cost = sum(`Total Cost`)/sum(`Total Volume of Bookings`),
            ROA = sum(Amount)/ sum(`Total Cost`),
            sum_total_cost = sum(`Total Cost`),
            sum_total_amount = sum(Amount),
            sum_total_bookings = sum(`Total Volume of Bookings`))

# check summary statictic of publishers
summary(group_by_pub)

# normalize every observation in the df using the for loop
group_by_pub$cpc_norm <- c()
group_by_pub$ctr_norm <-  c()
group_by_pub$tbpc_norm <- c()
group_by_pub$cac_norm <- c()
group_by_pub$roa_norm <- c()

# Use for loop to normalize for every observation
for (i in 1:nrow(group_by_pub)){
  group_by_pub$cpc_norm[i] <-  normalize(group_by_pub$cost_per_click[i], min(group_by_pub$cost_per_click), max(group_by_pub$cost_per_click))
  group_by_pub$ctr_norm[i] <-  normalize(group_by_pub$click_thru_rate[i], min(group_by_pub$click_thru_rate), max(group_by_pub$click_thru_rate))
  group_by_pub$tbpc_norm[i] <- normalize(group_by_pub$total_booking_per_click[i], min(group_by_pub$total_booking_per_click), max(group_by_pub$total_booking_per_click))
  group_by_pub$cac_norm[i] <-  normalize(group_by_pub$cus_acq_cost[i], min(group_by_pub$cus_acq_cost), max(group_by_pub$cus_acq_cost))
  group_by_pub$roa_norm[i] <-  normalize(group_by_pub$ROA[i], min(group_by_pub$ROA), max(group_by_pub$ROA))
}

# We defined a metric for each of our valuable business cases with weight
# score is the evaluation of our team matrices to each business case
group_by_pub$score <- c()
for (i in 1:nrow(group_by_pub)){
  group_by_pub$score[i] <-  -0.2 * group_by_pub$cpc_norm [i]+
    0.2 * group_by_pub$ctr_norm [i]+
    0.25 * group_by_pub$tbpc_norm[i] +
    -0.1 * group_by_pub$cac_norm [i]+
    0.25 * group_by_pub$roa_norm[i]
} 

# display in a histogram for the total score
hist(group_by_pub$score)


## plot 1 << Publisher Efficiency before optimize
ggplot(data=group_by_pub, aes(x=score, y=ROA)) +
  geom_point(aes(x=score, 
                 y=ROA, 
                 color=`Publisher Name`,
                 shape=`Publisher Name`,
                 size = sum_total_cost)) + 
  geom_rect(aes(xmin = -0.15, xmax = 0.46, ymin = -Inf, ymax = Inf),
            fill = "red", alpha = 0.02) +
  labs(title="Publisher Efficiency Overview", x = "Score", y= "ROA")  +
  scale_size_continuous("Total Cost", labels = scales::comma) + 
  theme(axis.text = element_text(size = 12),
        axis.title.x = element_text(size = 16),
        axis.title.y = element_text(size = 16)) +   
  scale_shape_manual(values = c(15,16,17,18,19,20,15)) +
  guides(color=guide_legend(override.aes = list(size = 4)))

# the mean is 0.46, so we define 0.46 as business success.
# we cut off any campaign which are under 0.46 score.
mean(group_by_pub$score)

# save publisher efficiency plot to local folder  
ggsave(file="Publisher Efficiency Overview.png",width=1850 ,height=950, units ="px", scale = 1.2 )

# group by publisher name and campaign
group_by_pub_cam <- af_df %>% 
  group_by(`Publisher Name`, Campaign) %>% 
  summarise(cost_per_click = sum(`Click Charges`) / sum(`Clicks`),
            click_thru_rate = round((sum(`Clicks`)/ sum(Impressions) ) * 100, 2),
            total_booking_per_click = round((sum(`Total Volume of Bookings`) / sum(Clicks))*100,2),
            cus_acq_cost = sum(`Total Cost`)/(sum(`Total Volume of Bookings`)+0.001),
            ROA = sum(Amount)/ sum(`Total Cost`),
            sum_total_cost = sum(`Total Cost`),
            sum_total_amount = sum(Amount),
            sum_total_bookings = sum(`Total Volume of Bookings`))

# check summary statistic of campaigns
summary(group_by_pub_cam)


# normalize every observation in the df using the for loop
group_by_pub_cam$cpc_norm <- c()
group_by_pub_cam$ctr_norm <-  c()
group_by_pub_cam$tbpc_norm <- c()
group_by_pub_cam$cac_norm <- c()
group_by_pub_cam$roa_norm <- c()

# Use for loop to normalize for every observation
for (i in 1:nrow(group_by_pub_cam)){
  group_by_pub_cam$cpc_norm[i]  <-  normalize(group_by_pub_cam$cost_per_click[i], min(group_by_pub_cam$cost_per_click), max(group_by_pub_cam$cost_per_click))
  group_by_pub_cam$ctr_norm[i]  <-  normalize(group_by_pub_cam$click_thru_rate[i], min(group_by_pub_cam$click_thru_rate), max(group_by_pub_cam$click_thru_rate))
  group_by_pub_cam$tbpc_norm[i] <-  normalize(group_by_pub_cam$total_booking_per_click[i], min(group_by_pub_cam$total_booking_per_click), max(group_by_pub_cam$total_booking_per_click))
  group_by_pub_cam$cac_norm[i]  <-  normalize(group_by_pub_cam$cus_acq_cost[i], min(group_by_pub_cam$cus_acq_cost), max(group_by_pub_cam$cus_acq_cost))
  group_by_pub_cam$roa_norm[i]  <-  normalize(group_by_pub_cam$ROA[i], min(group_by_pub_cam$ROA), max(group_by_pub_cam$ROA))
}

# We redefine our weight of each five business metric toward our score.
group_by_pub_cam$score <- c()
for (i in 1:nrow(group_by_pub_cam)){
  group_by_pub_cam$score[i] <- -0.2 * group_by_pub_cam$cpc_norm [i]+
    0.2 * group_by_pub_cam$ctr_norm [i]+
    0.25 * group_by_pub_cam$tbpc_norm[i] +
    -0.1 * group_by_pub_cam$cac_norm [i]+
    0.25 * group_by_pub_cam$roa_norm[i]
} 

# plot 2 - Campaign Efficiency before optimization
# ### Note: please maximize full screen to view the plot
ggplot(data=group_by_pub_cam, aes(x=score, y=ROA)) +
  geom_point(aes(x=score, 
                 y=ROA, 
                 color=Campaign,
                 size = sum_total_cost)) + 
  labs(x = "Score", y= "ROA") + 
  geom_vline(xintercept=c(-0.0765,0.0119), linetype="dashed", color = "red", linewidth = 0.75) +
  labs(title="Campaign Efficiency Overview", x = "Score", y= "ROA")  +
  scale_size_continuous("Total Cost", 
                        labels = scales::comma) + 
  theme(legend.key.size = unit(0.2,'cm'), 
        legend.justification=c(1,0), 
        legend.position=c(1,0), 
        legend.direction = "horizontal", 
        legend.box = "horizontal",
        legend.text = element_text(size = 12),
        axis.text = element_text(size = 12),
        axis.title.x = element_text(size = 16),
        axis.title.y = element_text(size = 16)) + 
  guides(color=guide_legend(ncol=2, override.aes = list(size = 4)), size=guide_legend(ncol=1)) +
  facet_wrap(~ `Publisher Name`) +  
  annotate(geom="text", x=-0.265, y=20, label="To \ndeactivate",
           color="red") +
  annotate(geom="text", x=-0.03, y=18, label="To\noptimize",
           color="red", ) +
  annotate(geom="text", x=0.2, y=20, label="To maintain",
           color="blue") +
  xlim(-0.3, 0.45)
ggsave(file="Campaign Efficiency Overview.png",width=1850 ,height=950, units ="px", scale = 3 )
####### Data Exploration and Visualization Ends ###############

############ Keyword Optimization Begin ######################

# Function to calculate group by publisher
group_data <- function(df, group_by_params) {
  group_by_pub <- df %>%
    group_by(across(all_of(group_by_params))) %>% 
    summarise(cost_per_click = round(sum(`Click Charges`) / sum(`Clicks`),2),
              click_thru_rate = round((sum(`Clicks`)/ sum(Impressions) ) * 100, 2),
              total_booking_per_click = round((sum(`Total Volume of Bookings`) / sum(Clicks))*100,2),
              # add 0.001 to treat Inf values by dividing 0
              cus_acq_cost = round(sum(`Total Cost`)/(sum(`Total Volume of Bookings`)+0.0001),2),
              ROA = round(sum(Amount)/ sum(`Total Cost`),2),
              sum_total_cost = round(sum(`Total Cost`)))
  return(group_by_pub)
}

# Function to converting binary for each business case and 
# converting all binary columns into numeric
calculate_binary_weights <- function(af_df) {
  af_df$cpc_norm <- numeric(nrow(af_df))
  af_df$ctr_norm <- numeric(nrow(af_df))
  af_df$tbpc_norm <- numeric(nrow(af_df))
  af_df$cac_norm <- numeric(nrow(af_df))
  af_df$roa_norm <- numeric(nrow(af_df))
  af_df$score <- numeric(nrow(af_df))
  # Use for loop to normalize for every observation
  for (i in 1:nrow(af_df)){
    af_df$cpc_norm[i]  <-  normalize(af_df$cost_per_click[i], min(af_df$cost_per_click, na.rm = TRUE), max(af_df$cost_per_click, na.rm = TRUE))
    af_df$ctr_norm[i]  <-  normalize(af_df$click_thru_rate[i], min(af_df$click_thru_rate, na.rm = TRUE), max(af_df$click_thru_rate, na.rm = TRUE))
    af_df$tbpc_norm[i] <-  normalize(af_df$total_booking_per_click[i], min(af_df$total_booking_per_click[is.finite(af_df$total_booking_per_click)]), max(af_df$total_booking_per_click[is.finite(af_df$total_booking_per_click)]))
    af_df$cac_norm[i]  <-  normalize(af_df$cus_acq_cost[i], min(af_df$cus_acq_cost, na.rm = TRUE), max(af_df$cus_acq_cost, na.rm = TRUE))
    af_df$roa_norm[i]  <-  normalize(af_df$ROA[i], min(af_df$ROA[is.finite(af_df$ROA)], na.rm = TRUE), max(af_df$ROA[is.finite(af_df$ROA)], na.rm = TRUE))
  }
  # Calculate total binary weight
  for (i in 1:nrow(af_df)){
    af_df$score[i] <- -0.2 * af_df$cpc_norm [i]+
                       0.2 * af_df$ctr_norm [i]+
                       0.25 * af_df$tbpc_norm[i]-
                       0.1 * af_df$cac_norm [i]+
                       0.25 * af_df$roa_norm[i]
  }
  # return data frame
  return(af_df)
}


# Function to convert binary to numeric
optmize <- function(df) {
# Code for converting binary for each business case and converting all binary columns into numeric
  df$cpc_binary <- df$cost_per_click <= mean(df$cost_per_click, na.rm = TRUE)
  df$cpc_binary <- as.numeric(df$cpc_binary)
  return (df)
  }

# Use a new data frame to optimize the keyword
af_df <- read_xls(path=file_path,
                         sheet=2,
                         col_names = TRUE)

# Calculate group by publisher
group_by_pub_campaign <- group_data(af_df,c("Publisher Name", "Campaign"))
group_by_pub_campaign_keywords <- group_data(af_df,c("Publisher Name", "Campaign", "Keyword", "Keyword ID"))
group_by_pub <- group_data(af_df,c("Publisher Name"))

#
# Calculate binary weights
group_by_pub_campaign_with_weights <- calculate_binary_weights(group_by_pub_campaign)
group_by_pub_campaign_keywords_with_weights <- calculate_binary_weights(group_by_pub_campaign_keywords)

# check summary statistic of group by column
summary(group_by_pub_campaign_with_weights$score)
summary(group_by_pub_campaign_with_weights[group_by_pub_campaign_with_weights$`Publisher Name`=='Yahoo - US',]$score)
summary(group_by_pub_campaign_with_weights$score[])

#Q3 and Yahoo
# Function to subset data frame to preapre the optimization of Quantile 3 and Yahoo campaign
filter_data <- function(df, df2) {
  indices <- which(
    (df2$score > -0.076542 & df2$score < 0.011998) | (df2$`Publisher Name` == 'Yahoo - US'&df2$score > -0.05046)
      )
  
  filtered_df <- df2[indices, ]
  filtered_af_df <- merge(df, filtered_df[, c("Publisher Name", "Campaign")], by = c("Publisher Name", "Campaign"))
  
  return(filtered_af_df)
}

# call Q3 and yahoo optimize function
targetForOptmize<- filter_data(group_by_pub_campaign_keywords_with_weights,group_by_pub_campaign_with_weights)

# check summary statistic for optimized target
summary(targetForOptmize$score)

# Function to subset data frame which is under score 0.00593
filter_data_1 <- function(df, df2) {
  indices <- which(
    (df2$score < 0.00593)
  )
  
  filtered_df <- df2[indices, ]
  filtered_af_df <- df[!df$`Keyword ID` %in% filtered_df$`Keyword ID`,]
  
  return(filtered_af_df)
}

# call the score under 0.00593 function to remove keywords and export excel.
data_RemovedKeyword <- filter_data_1(af_df,targetForOptmize)
data_RemovedKeyword_by_pub_campaign <- group_data(data_RemovedKeyword,c("Publisher Name", "Campaign"))
data_RemovedKeyword_by_pub_campaign_with_weight <- calculate_binary_weights(data_RemovedKeyword_by_pub_campaign)


#Q1,2,3,Yahoo
# Function to subset yahoo campaign
filter_data <- function(df, df2) {
  indices <- which(
    (df2$score > -0.076542) | (df2$`Publisher Name` == 'Yahoo - US'&df2$score > -0.05046)
  )
  filtered_df <- df2[indices, ]
  filtered_af_df <- merge(df, filtered_df[, c("Publisher Name", "Campaign")], by = c("Publisher Name", "Campaign"))
  return(filtered_af_df)
}

# call the yahoo optimize function and store in a new variable
targetForOptmize<- filter_data(group_by_pub_campaign_keywords_with_weights,group_by_pub_campaign_with_weights)

# check the summary statistic for optimized yahoo campaign
summary(targetForOptmize$score)

# Function to subset keywords with bad performance
filter_data_1 <- function(df, df2) {
  indices <- which(
    (df2$score < -0.019753)
  )
  filtered_df <- df2[indices, ]
  filtered_af_df <- df[!df$`Keyword ID` %in% filtered_df$`Keyword ID`,]
  return(filtered_af_df)
}

# remove bad performance keyword using defined function
data_RemovedKeyword <- filter_data_1(af_df,targetForOptmize)


# Campaing level
# check summary statistic for campaign
summary(group_by_pub_campaign)

# function for calcuating binary weigh of campaign
calculate_binary_weights_cam <- function(af_df) {
  # Code for converting binary for each business case and converting all binary columns into numeric
  af_df$cpc_norm <- numeric(nrow(af_df))
  af_df$ctr_norm <- numeric(nrow(af_df))
  af_df$tbpc_norm <- numeric(nrow(af_df))
  af_df$cac_norm <- numeric(nrow(af_df))
  af_df$roa_norm <- numeric(nrow(af_df))
  af_df$score <- numeric(nrow(af_df))
  # Use for loop to normalize for every observation
  for (i in 1:nrow(af_df)){
    af_df$cpc_norm[i]  <-  normalize(af_df$cost_per_click[i], 0.350, max(af_df$cost_per_click, na.rm = TRUE))
    af_df$ctr_norm[i]  <-  normalize(af_df$click_thru_rate[i], 0.340, max(af_df$click_thru_rate, na.rm = TRUE))
    af_df$tbpc_norm[i] <-  normalize(af_df$total_booking_per_click[i], 0, max(af_df$total_booking_per_click[is.finite(af_df$total_booking_per_click)]))
    af_df$cac_norm[i]  <-  normalize(af_df$cus_acq_cost[i], 46, max(af_df$cus_acq_cost, na.rm = TRUE))
    af_df$roa_norm[i]  <-  normalize(af_df$ROA[i], 0, max(af_df$ROA[is.finite(af_df$ROA)], na.rm = TRUE))
  }
  # Calculate total binary weight
  for (i in 1:nrow(af_df)){
    af_df$score[i] <- -0.2 * af_df$cpc_norm [i]+
      0.2 * af_df$ctr_norm [i]+
      0.25 * af_df$tbpc_norm[i] +
      -0.1 * af_df$cac_norm [i]+
      0.25 * af_df$roa_norm[i]
  }
  return(af_df)
}

# Call group data function to group optimized keyword campaign
data_RemovedKeyword_by_pub_campaign <- group_data(data_RemovedKeyword,c("Publisher Name", "Campaign"))
data_RemovedKeyword_by_pub_campaign_with_weights <- calculate_binary_weights_cam(data_RemovedKeyword_by_pub_campaign)

# pub level
# check summary statistic of optimized group by campaign data frame
summary(group_by_pub)

# Function for calculating weight of keyword performance for publisher
calculate_binary_weights_pub <- function(af_df) {
  # Code for converting binary for each business case and converting all binary columns into numeric
  af_df$cpc_norm <- numeric(nrow(af_df))
  af_df$ctr_norm <- numeric(nrow(af_df))
  af_df$tbpc_norm <- numeric(nrow(af_df))
  af_df$cac_norm <- numeric(nrow(af_df))
  af_df$roa_norm <- numeric(nrow(af_df))
  af_df$score <- numeric(nrow(af_df))
  # Use for loop to normalize for every observation
  for (i in 1:nrow(af_df)){
    af_df$cpc_norm[i]  <-  normalize(af_df$cost_per_click[i], 1.013, max(af_df$cost_per_click, na.rm = TRUE))
    af_df$ctr_norm[i]  <-  normalize(af_df$click_thru_rate[i], 0.340, max(af_df$click_thru_rate, na.rm = TRUE))
    af_df$tbpc_norm[i] <-  normalize(af_df$total_booking_per_click[i], 0.240, max(af_df$total_booking_per_click[is.finite(af_df$total_booking_per_click)]))
    af_df$cac_norm[i]  <-  normalize(af_df$cus_acq_cost[i], 69.79, max(af_df$cus_acq_cost, na.rm = TRUE))
    af_df$roa_norm[i]  <-  normalize(af_df$ROA[i], 2.447, max(af_df$ROA[is.finite(af_df$ROA)], na.rm = TRUE))
  }
  # Calculate total binary weight
  for (i in 1:nrow(af_df)){
    af_df$score[i] <- -0.2 * af_df$cpc_norm [i]+
      0.2 * af_df$ctr_norm [i]+
      0.25 * af_df$tbpc_norm[i] +
      -0.1 * af_df$cac_norm [i]+
      0.25 * af_df$roa_norm[i]
  }
  return(af_df)
}

# call group by function to group by publisher with optimized keyword
data_RemovedKeyword_by_pub <- group_data(data_RemovedKeyword,c("Publisher Name"))
data_RemovedKeyword_by_pub_with_weight <- calculate_binary_weights_pub(data_RemovedKeyword_by_pub)

# Call group by function to group by publisher from original data frame to compare
# the optimization performance
group_by_pub <- group_data(af_df,c("Publisher Name"))
############ Keyword Optimization Ends ######################

############ Comparison with Kayak Begin ######################
# Set the file name and path
file_name <- "Comparison with Kayak.png"
file_path <- "./"

# Create the graphics device
png(paste0(file_path, file_name), width = 800, height = 600)
vs_kayak <- data.frame(
  cost_per_click = c(0.87999, 1.25642832),
  total_booking_per_click = c(1.21429, 7.3265234),
  cus_acq_cost = c(81.12558, 17.14903846),
  ROA = c(17.34649, 65.51555929),
  row.names = c("PublishersAVG", "Kayak")
)

# Define values
# in our calculation we multiple cost per click with 100
# So, we are multiplying kayak cost per click with 100 as well
vs_kayak <- data.frame( cost_per_click = c(0.87999, 1.25642832),
                        total_booking_per_click = c(1.21429, 7.3265234),
                        cus_acq_cost = c(81.12558, 17.14903846),
                        ROA = c(17.34649, 65.51555929),
                        row.names = c("PublishersAVG", "Kayak"))

# selecting colors
my_colors <- c("#a5c1d5", "#4ba1e2", "#547c8c", "#214e7a")

# Create the barplot
barplot(t(vs_kayak), beside = TRUE, col = my_colors,
            xlab = "Variables", ylab = "Values", main = "Comparison against Kayak")

# Add the legend for the colors
legend("topright", inset = 0.25, legend = names(vs_kayak),
       fill = my_colors, cex=1)

# Close the graphics device and save the plot to a file
dev.off()
############ Comparison with Kayak Ends ######################

################ Summary ######################
# 1. Plot 1 - Publisher Efficiency Plot (Before Optimization)
#    Saved to local folder under the file name "Publisher Efficiency Overview.png"

# 2. Plot 2 - Campaign Efficiency Plot (Before Optimization)
#    Saved to local folder under the file name "Campaign Efficiency Overview.png"

# 3A. Publisher Metrics Before Optimization
      View(group_by_pub)
# 3B. Publisher Metrics After Optimization
      View(data_RemovedKeyword_by_pub)

# 4. Comparison with Kayak
#     Save to local folder under the file name "Comparison with Kayak.png"
      