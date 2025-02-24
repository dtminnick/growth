library("dplyr")
library("openxlsx")

project_transactions <- function(baseline, baseline_digital, growth_rate, shift_rate, work_effort, number_weeks) {
  
  time <- 0:number_weeks
  
  e <- 2.718
  
  # Manual transactions projection
  
  transactions_remaining <- round(baseline * e^((growth_rate + shift_rate) * time), 0)
  
  reduction <- -round(c(0, diff(transactions_remaining)), 0)
  
  # Full conversion of manual reductions to digital
  
  digital_transactions <- baseline_digital + cumsum(reduction)
  
  # Ensure digital transactions never go below initial baseline
  
  digital_transactions <- pmax(digital_transactions, baseline_digital)
  
  # Track digital transaction increases per week
  
  digital_increase <- c(0, diff(digital_transactions))
  
  # Effort saved calculations
  
  effort_saved_hours <- round(reduction * work_effort, 2)
  
  effort_saved_fte <- round(effort_saved_hours / 1456, 2)
  
  cumulative_fte <- round(cumsum(effort_saved_fte), 2)
  
  return(data.frame(week = time,
                    transactions_remaining = transactions_remaining,
                    transaction_reduction = reduction,
                    digital_transactions = digital_transactions,
                    digital_increase = digital_increase,
                    effort_saved_hours = effort_saved_hours,
                    effort_saved_fte = effort_saved_fte,
                    cumulative_fte_saved = cumulative_fte))
}


# Read input data
input_data <- read.csv("./inst/extdata/test_input_data.txt") %>%
  mutate(growth_rate = as.numeric(growth_rate),
         shift_rate = as.numeric(shift_rate),
         digital_growth_rate = as.numeric(digital_growth_rate))  # New column for digital transaction growth rate

# Apply the model to each row in the input data
tables <- lapply(1:nrow(input_data), function(i) {
  
  row <- input_data[i, ]
  
  df <- project_transactions(row$baseline, row$baseline_digital, row$growth_rate, row$shift_rate, 
                             row$digital_growth_rate, row$work_effort, number_weeks = 52)
  
  df$work_type <- row$work_type
  df$completion_type <- row$completion_type
  df$channel <- row$channel
  
  return(df)
})

# Combine results
all_data <- bind_rows(tables) %>%
  group_by(work_type, week, channel) %>%
  summarise(transactions_remaining = sum(transactions_remaining),
            transaction_reduction = sum(transaction_reduction),
            digital_transactions = sum(digital_transactions),
            digital_increase = sum(digital_increase),
            effort_saved_hours = sum(effort_saved_hours),
            effort_saved_fte = sum(effort_saved_fte),
            cumulative_fte_saved = sum(cumulative_fte_saved))

# Summary aggregation
summary_data <- all_data %>%
  group_by(week, channel) %>%
  summarise(transactions_remaining = sum(transactions_remaining),
            transaction_reduction = sum(transaction_reduction),
            digital_transactions = sum(digital_transactions),
            digital_increase = sum(digital_increase),
            effort_saved_hours = sum(effort_saved_hours),
            effort_saved_fte = sum(effort_saved_fte),
            cumulative_fte_saved = sum(cumulative_fte_saved))

# Export to Excel
wb <- createWorkbook()

left_align <- createStyle(halign = "left")
right_align <- createStyle(halign = "right")

for (i in 1:length(tables)) {
  
  sheet_name <- paste(tables[[i]][1, "work_type"], tables[[i]][2, "completion_type"], tables[[i]][3, "channel"], sep = "-")
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet = sheet_name, tables[[i]])
  setColWidths(wb, sheet_name, cols = 1:10, widths = rep(20, 10))
  
  addStyle(wb, sheet_name, style = left_align, cols = 1:3, rows = 1:nrow(tables[[i]]) + 1, gridExpand = TRUE)
  addStyle(wb, sheet_name, style = right_align, cols = 4:10, rows = 1:nrow(tables[[i]]) + 1, gridExpand = TRUE)
}

addWorksheet(wb, "All Combined")
writeData(wb, "All Combined", summary_data)
setColWidths(wb, "All Combined", cols = 1:9, widths = rep(20, 9))

addStyle(wb, "All Combined", style = left_align, cols = 2, rows = 1:nrow(summary_data) + 1, gridExpand = TRUE)
addStyle(wb, "All Combined", style = right_align, cols = 3:9, rows = 1:nrow(summary_data) + 1, gridExpand = TRUE)

saveWorkbook(wb, "./output/test_output_file.xlsx", overwrite = TRUE)
