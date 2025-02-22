
library("dplyr")
library("openxlsx")

project_transactions <- function(baseline, growth_rate, shift_rate, work_effort, number_weeks) {
  
  time <- 0:number_weeks
  
  e <- 2.718
  
  transactions_remaining <- round(baseline * e^((growth_rate - shift_rate) * time), 0)
  
  reduction <- -round(c(0, diff(transactions_remaining)), 0)
  
  effort_saved_hours <- round(reduction * work_effort, 2)
  
  effort_saved_fte <- round(effort_saved_hours / 1456, 2)
  
  cumulative_fte <- round(cumsum(effort_saved_fte), 2)
  
  return(data.frame(week = time,
                    transactions_remaining = transactions_remaining,
                    transaction_reduction = reduction,
                    effort_saved_hours = effort_saved_hours,
                    effort_saved_fte = effort_saved_fte,
                    cumulative_fte_saved = cumulative_fte))

}

input_data <- read.csv("./inst/extdata/test_input_data.txt")

tables <- lapply(1:nrow(input_data), function(i) {
  
  row <- input_data[i, ]
  
  df <- project_transactions(row$baseline, row$growth_rate, row$shift_rate, row$work_effort, number_weeks = 52)
  
  df$work_type <- row$work_type
  
  df$completion_type <- row$completion_type
  
  df <- df[, c("work_type", "completion_type", 
               setdiff(names(df), c("work_type", "completion_type")))]
  
  return(df)
  
})

all_data <- bind_rows(tables) %>%
  select(work_type,
         week,
         transactions_remaining,
         transaction_reduction,
         effort_saved_hours,
         effort_saved_fte,
         cumulative_fte_saved) %>%
  group_by(work_type, week) %>%
  summarise(transactions_remaining = sum(transactions_remaining),
            transaction_reduction = sum(transaction_reduction),
            effort_saved_hours = sum(effort_saved_hours),
            effort_saved_fte = sum(effort_saved_fte),
            cumulative_fte_saved = sum(cumulative_fte_saved))

summary_data <- all_data %>%
  group_by(week) %>%
  summarise(transactions_remaining = sum(transactions_remaining),
            transaction_reduction = sum(transaction_reduction),
            effort_saved_hours = sum(effort_saved_hours),
            effort_saved_fte = sum(effort_saved_fte),
            cumulative_fte_saved = sum(cumulative_fte_saved))

wb <- createWorkbook()

left_align  <- createStyle(halign = "left")

right_align <- createStyle(halign = "right")

center_align <- createStyle(halign = "center")

for (i in 1:length(tables)) {
  
  num_rows <- nrow(tables[[i]]) + 1
  
  sheet_name <- paste(tables[[i]][1, "work_type"],
                      tables[[i]][2, "completion_type"],
                      sep = "-")
  
  addWorksheet(wb, sheet_name)
  
  writeData(wb, sheet = sheet_name, tables[[i]])
  
  setColWidths(wb, sheet_name, cols = 1:8, widths = c(20, 20, 20, 20, 20, 20, 20, 20)) 
  
  addStyle(wb, sheet_name, 
           style = left_align,  
           cols = 1:2, 
           rows = 1:num_rows, 
           gridExpand = TRUE)
  
  addStyle(wb, 
           sheet_name, 
           style = right_align, 
           cols = 3:8, 
           rows = 1:num_rows, 
           gridExpand = TRUE)
  
}

addWorksheet(wb, "All Combined")

writeData(wb, "All Combined", summary_data)

setColWidths(wb, "All Combined", cols = 1:6, widths = c(20, 20, 20, 20, 20, 20))

addStyle(wb, "All Combined", 
         style = left_align,  
         cols = 1, 
         rows = 1:num_rows, 
         gridExpand = TRUE)

addStyle(wb, 
         "All Combined", 
         style = right_align, 
         cols = 2:6, 
         rows = 1:num_rows, 
         gridExpand = TRUE)

saveWorkbook(wb, "./output/test_output_file.xlsx", overwrite = TRUE)
