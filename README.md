# Excel Data Explorer (Shiny)

A lightweight Shiny app to explore datasets from Excel files:
- Upload an `.xlsx` / `.xls`
- Choose the sheet
- Apply **multiple filters** (categorical + numeric range)
- Create **bar plots**, **pie charts**, **scatter plots**, and **line charts**
- Customize colors using palettes or manual level-to-color mapping

## Demo data
Example dataset is available in `data/test_data.xlsx`.

## Quick demo
1. Run the app
2. Upload `data/test_data.xlsx`
3. Choose a sheet
4. Add a filter (e.g. smoker = no)
5. Make a barplot on `job` or a scatter `salary vs debth` (if present)

## Requirements
Install packages in R:

```r
install.packages(c("shiny","readxl","dplyr","tidyr","ggplot2","stringr"))


