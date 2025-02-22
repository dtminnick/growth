
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
  
  df <- df[, c("work_type", setdiff(names(df), "work_type"))] 
  
  df
  
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

for (i in 1:length(tables)) {
  
  sheet_name <- paste(tables[[i]][1, "work_type"],
                      tables[[i]][2, "completion_type"],
                      sep = "-")
  
  addWorksheet(wb, sheet_name)
  
  writeData(wb, sheet = sheet_name, tables[[i]])
  
}

addWorksheet(wb, "All Combined")

writeData(wb, "All Combined", summary_data)

saveWorkbook(wb, "my_excel_file.xlsx", overwrite = TRUE)
