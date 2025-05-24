# Growth Modeling: Manual-to-Digital Transaction Shift

This R project models the projected shift from manual to digital transactions over time, using configurable input parameters. It calculates weekly transaction volumes, digital conversion, and estimated effort savings (in hours and FTE), and exports formatted results to Excel for analysis.

# Key Features
- Models transactional growth and channel shift dynamics over time.
- Estimates effort savings based on task-level work effort.
- Generates individualized and summary reports by work type, completion type, and channel.
- Outputs clean, formatted Excel workbooks with aligned styles.

# Use Case
This tool is designed to support operations leaders and continuous improvement professionals who need to:
- Forecast the impact of digital adoption initiatives,
- Quantify operational benefits (hours/FTEs saved), and
- Report outcomes in a business-friendly Excel format.

# Input Format
The model expects a `.txt` file with input columns:
- `baseline`
- `baseline_digital`
- `growth_rate`
- `shift_rate`
- `digital_growth_rate`
- `work_effort`
- `work_type`, `completion_type`, `channel`

Example file: `./inst/extdata/test_input_data.txt`

# Output
The model produces:
- One worksheet per row in the input data
- A combined summary worksheet (`All Combined`)
- Styled and formatted Excel file at `./output/test_output_file.xlsx`

*Built as part of my data science journey to apply analytics in service of business transformation and operational efficiency.*
