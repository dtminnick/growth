
library("dplyr")
library("openxlsx")

project_transactions <- function(baseline_manual, 
                                 baseline_digital, 
                                 growth_rate, 
                                 shift_rate, 
                                 work_effort, 
                                 number_weeks, 
                                 start_digital_shift) {
  
  time <- 0:number_weeks
  
  e <- 2.718
  
  before_shift <- time < start_digital_shift
  
  manual_remaining <- ifelse(before_shift,
                             round(baseline_manual * e^(growth_rate * time), 0),
                             round(baseline_manual * e^((shift_rate - growth_rate) * time), 0))
  
  manual_reduction <- -round(c(0, diff(manual_remaining)), 0)
  
  digital_transactions <- ifelse(before_shift,
                                 round(baseline_digital * e^(growth_rate * time), 0),
                                 round(baseline_digital * e^(growth_rate * time) + cumsum(abs(manual_reduction)) * (time >= start_digital_shift), 0))
  
  digital_increase <- c(0, diff(digital_transactions))
  
  effort_saved_hours <- round(manual_reduction * work_effort, 2)
  
  effort_saved_fte <- round(effort_saved_hours / 1456, 2)
  
  cumulative_fte <- round(cumsum(effort_saved_fte), 2)
  
  return(data.frame(week = time,
                    manual_remaining = manual_remaining,
                    manual_reduction = manual_reduction,
                    digital_transactions = digital_transactions,
                    digital_increase = digital_increase,
                    effort_saved_hours = effort_saved_hours,
                    effort_saved_fte = effort_saved_fte,
                    cumulative_fte_saved = cumulative_fte))

}

input_data <- read.csv("./inst/extdata/test_input_data.txt") %>%
  mutate(growth_rate = as.numeric(growth_rate),
         shift_rate = as.numeric(shift_rate))

tables <- lapply(1:nrow(input_data), function(i) {
  
  row <- input_data[i, ]
  
  df <- project_transactions(row$baseline_manual,
                             row$baseline_digital,
                             row$growth_rate, 
                             row$shift_rate, 
                             row$work_effort, 
                             row$number_weeks,
                             row$start_digital_shift)
  
  df$work_type <- row$work_type
  
  df$completion_type <- row$completion_type
  
  df <- df[, c("work_type", "completion_type", 
               setdiff(names(df), c("work_type", "completion_type")))]
  
  return(df)
  
})

all_data <- bind_rows(tables) %>%
  select(work_type,
         week,
         manual_remaining,
         manual_reduction,
         digital_transactions,
         digital_increase,
         effort_saved_hours,
         effort_saved_fte,
         cumulative_fte_saved) %>%
  group_by(work_type, week) %>%
  summarise(manual_remaining = sum(manual_remaining),
            manual_reduction = sum(manual_reduction),
            digital_transactions = sum(digital_transactions),
            digital_increase = sum(digital_increase),
            effort_saved_hours = sum(effort_saved_hours),
            effort_saved_fte = sum(effort_saved_fte),
            cumulative_fte_saved = sum(cumulative_fte_saved))

summary_data <- all_data %>%
  group_by(week) %>%
  summarise(manual_remaining = sum(manual_remaining),
            manual_reduction = sum(manual_reduction),
            digital_transactions = sum(digital_transactions),
            digital_increase = sum(digital_increase),
            effort_saved_hours = sum(effort_saved_hours),
            effort_saved_fte = sum(effort_saved_fte),
            cumulative_fte_saved = sum(cumulative_fte_saved))

wb <- createWorkbook()

left_align  <- createStyle(halign = "left")

right_align <- createStyle(halign = "right")

center_align <- createStyle(halign = "center")

number_format <- createStyle(numFmt = "#,##0")

decimal_format <- createStyle(numFmt = "#,##0.00")

for (i in 1:length(tables)) {
  
  num_rows <- nrow(tables[[i]]) + 1
  
  sheet_name <- paste(tables[[i]][1, "work_type"],
                      tables[[i]][2, "completion_type"],
                      sep = "-")
  
  addWorksheet(wb, sheet_name)
  
  writeData(wb, sheet = sheet_name, tables[[i]])
  
  setColWidths(wb, sheet_name, cols = 1:9, widths = c(20, 20, 20, 20, 20, 20, 20, 20)) 
  
  addStyle(wb, sheet_name, 
           style = left_align,  
           cols = 1:3, 
           rows = 1:num_rows, 
           gridExpand = TRUE)
  
  addStyle(wb, 
           sheet_name, 
           style = right_align, 
           cols = 4:8, 
           rows = 1:num_rows, 
           gridExpand = TRUE)
  
  addStyle(wb, 
           sheet_name, 
           style = number_format, 
           cols = 4:6, 
           rows = 2:num_rows, 
           gridExpand = TRUE)
  
  addStyle(wb, 
           sheet_name, 
           style = decimal_format, 
           cols = 7:8, 
           rows = 2:num_rows, 
           gridExpand = TRUE)
  
  
}

addWorksheet(wb, "All Combined")

writeData(wb, "All Combined", summary_data)

setColWidths(wb, "All Combined", cols = 1:7, widths = c(20, 20, 20, 20, 20, 20, 20))

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

addStyle(wb, 
         "All Combined", 
         style = number_format, 
         cols = 2:4, 
         rows = 2:num_rows, 
         gridExpand = TRUE)

addStyle(wb, 
         "All Combined", 
         style = decimal_format, 
         cols = 5:6, 
         rows = 2:num_rows, 
         gridExpand = TRUE)

saveWorkbook(wb, "./output/test_output_file.xlsx", overwrite = TRUE)
