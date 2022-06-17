library(mailR)
library(tidyverse)
library(readr)
library(data.table)
library(odbc)
library(DBI)
library(dbplyr)
library(openxlsx)
library(lubridate)
library(stringi)
library(openxlsx)



setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
CrisisGo_No_Response_Data <- read_excel("CrisisGo No Response Data Ste 2021_09_07.xlsx")%>%
  arrange(`Supervisor 1 Email`)

unique_supervisors_1<- CrisisGo_No_Response_Data %>%
  select(`Supervisor 1 Email`, `Supervisor 1` ) %>%
  unique()

unique_supervisors_2<- CrisisGo_No_Response_Data %>%
  select(`Supervisor 2 Email`, `Supervisor 2`) %>%
  arrange(`Supervisor 2 Email`) %>%
  filter(!is.na(`Supervisor 2 Email`)) %>%
  unique()



subject <- "ACTION NEEED: COVID-19 Vaccination Disclosure"




for(i in 1:length(unique_supervisors_1$`Supervisor 1 Email`))
# for(i in 60:60)
{

  from <- "email"
  # to <-  'email'
  to <-unique_supervisors_1$`Supervisor 1 Email`[i]
  print(to)
  workbook_name <- "Employees with no response"
  save_location = str_interp("${workbook_name}.xlsx")

  filtered_data<- CrisisGo_No_Response_Data %>%
    filter(`Supervisor 1 Email` ==unique_supervisors_1$`Supervisor 1 Email`[i] )
 
    workbook<- createWorkbook()
    addWorksheet(workbook, workbook_name)
    writeDataTable(workbook, sheet = workbook_name,  x = filtered_data, startRow = 1, startCol = 1)
    setColWidths(workbook, sheet = workbook_name, cols = 1:9, widths = "auto")
    setRowHeights(workbook, sheet = workbook_name, rows = 1, heights = 40)
    addStyle(workbook, sheet = workbook_name, style = createStyle(wrapText = TRUE), rows = 1, cols = 1:9)
    print(saveWorkbook(workbook, save_location, overwrite = TRUE,  returnValue = TRUE))


  first_last_name<- unique_supervisors_1$`Supervisor 1`[i]
  print(first_last_name)

msg <- str_interp("Good Afternoon ${first_last_name},
                 
The Guilford County Schools Board of Education requires that all employees provide their vaccination status and proof of vaccination, if applicable, using the CrisiGo Vaccination Status Form. Please find the attached list of employees in your department who have not yet provided their vaccination status. The deadline for providing this information was Friday, September 3. T

Please contact these employees to remind them to complete the form immediately.

Thank you for your assistance.
")

  send.mail(from = from,
            to = to,
            subject = subject,
            body = msg,
            authenticate = TRUE,
            attach.files = "Employees with no response.xlsx",
            smtp = list(host.name = "smtp.office365.com", port = 587,
                        user.name = "email", passwd = "email*", tls = TRUE),
            send = TRUE,debug = FALSE, html = FALSE)

}


for(i in 1:length(unique_supervisors_2$`Supervisor 2 Email`))
# for(i in 11:11)
{
 
  from <- "email"
  # to <-  'email'
  to <-unique_supervisors_2$`Supervisor 2 Email`[i]
  print(to)
  workbook_name <- "Employees with no response"
  save_location = str_interp("${workbook_name}.xlsx")
 
  filtered_data<- CrisisGo_No_Response_Data %>%
    filter(`Supervisor 2 Email` ==unique_supervisors_2$`Supervisor 2 Email`[i] )
 
  workbook<- createWorkbook()
  addWorksheet(workbook, workbook_name)
  writeDataTable(workbook, sheet = workbook_name,  x = filtered_data, startRow = 1, startCol = 1)
  setColWidths(workbook, sheet = workbook_name, cols = 1:9, widths = "auto")
  setRowHeights(workbook, sheet = workbook_name, rows = 1, heights = 40)
  addStyle(workbook, sheet = workbook_name, style = createStyle(wrapText = TRUE), rows = 1, cols = 1:9)
  saved = saveWorkbook(workbook, save_location, overwrite = TRUE,  returnValue = TRUE)
 
 
  first_last_name<- unique_supervisors_2$`Supervisor 2`[i]
  print(first_last_name)
 
  msg <- str_interp("Good Afternoon ${first_last_name},
                 
The Guilford County Schools Board of Education requires that all employees provide their vaccination status and proof of vaccination, if applicable, using the CrisiGo Vaccination Status Form. Please find the attached list of employees in your department who have not yet provided their vaccination status. The deadline for providing this information was Friday, September 3. T

Please contact these employees to remind them to complete the form immediately.

Thank you for your assistance.
")
  send.mail(from = from,
            to = to,
            subject = subject,
            body = msg,
            authenticate = TRUE,
            attach.files = "Employees with no response.xlsx",
            smtp = list(host.name = "smtp.office365.com", port = 587,
                        user.name = "email", passwd = "password", tls = TRUE),
            send = TRUE,debug = FALSE, html = FALSE)
 
}