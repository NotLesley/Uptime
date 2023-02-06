#Author Lesley Chingwena
#Description: This script copies data from uptime csv file to the KIP report

import win32com.client as win32
import csv
from pywinauto.application import Application

sink = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Sensor_Uptime_Report_31_Jan 2023.xlsm"
source = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\uptime_2023_02_02.csv"

# Load the source workbook
excel = win32.gencache.EnsureDispatch('Excel.Application')
source_wb = excel.Workbooks.Add()

# Get the active sheet
source_sheet = source_wb.ActiveSheet

# Open CSV and read it
with open(source) as file:
    reader = csv.reader(file)
    data = [row for row in reader]

# Write the data to the sheet
for row in range(len(data)):
    for col in range(len(data[row])):
        source_sheet.Cells(row + 1, col + 1).Value = data[row][col]
        
# Load the target workbook
target_wb = excel.Workbooks.Open(sink, ReadOnly=False)

# Get the existing sheet "Metric" in the existing workbook
raw_data_sheet = target_wb.Sheets("data_raw")

# Clear the existing data in the "Metric" sheet
raw_data_sheet.Cells.ClearContents()

# Copy the data from the source sheet to the "Metric" sheet
source_sheet.Range("A1:Q735").Copy(raw_data_sheet.Range("A1"))

# Save the target workbook as .xlsm
target_wb.SaveAs("Sensor_Uptime_Report_05_Feb_2023.xlsm", FileFormat=52)

# Run the macro that corresponds to the button you want to press
excel.Application.Run("Sheet1.Copy New Uptime Data")

# Close the workbooks
source_wb.Close(SaveChanges=False)
target_wb.Close(SaveChanges=True)

# Quit Excel
excel.Quit()

print("ALL_Ok")




