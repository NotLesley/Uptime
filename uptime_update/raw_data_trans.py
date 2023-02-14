#Author Lesley Chingwena
#Description: This script copies data from uptime csv file to the KIP report "raw_data"
import os
import win32com.client as win32
import csv
import time

_DELAY = 100.0  # second
_QUIT_DELAY = 10.0

def update(source, weekly, daily, date, yesterday):

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
            
    # Load the weekly uptime workbook
    weekly_wb = excel.Workbooks.Open(weekly, ReadOnly=False)
    excel.Visible = False
    excel.DisplayAlerts=False

    # Get the existing sheet "Metric" in the existing workbook
    raw_data_sheet = weekly_wb.Sheets("data_raw")

    # Clear the existing data in the "Metric" sheet
    raw_data_sheet.Cells.ClearContents()

    # Copy the data from the source sheet to the "Metric" sheet
    source_sheet.Range("A1:Q735").Copy(raw_data_sheet.Range("A1"))

    # Close the workbooks
    source_wb.Close(SaveChanges=False)

    # Run the macro that corresponds to the VBA button (Copy New Uptime Data)
    excel.Run('Sheet1.CommandButton1_Click')

    # Delay to allow for KPI calculations before further manupulation
    time.sleep(_DELAY)

    daily_wb = excel.Workbooks.Open(daily, ReadOnly=False)
    
    # Update Daily KPI
    tdy_sheet = weekly_wb.Sheets(f"{date}")

    # Determine the size of data in the existing sheet
    data_range = tdy_sheet.UsedRange

    # Copy data from existing sheet to new sheet 
    tdy_sheet.Copy(Before=daily_wb.Sheets("VBA"))
    print("copy complete")

    new_sheet = daily_wb.Sheets(f"{date}")

    # Copy formatting from previous sheet 
    prev_sheet = daily_wb.Sheets(f"{yesterday}")
    s_range = data_range.Columns("S").Rows.Count
    prev_sheet.Range(f"S1:S{s_range}").Copy(new_sheet.Range("S1"))
    print("Formatting complete")

    # Update Metrics sheet 
    metric_sheet = daily_wb.Sheets("Metric")
    last_entry = metric_sheet.UsedRange.Rows.count
    
    # Copying formatting into the folowing row
    metric_sheet.Cells((last_entry + 1), 1).Value = metric_sheet.Cells(last_entry, 1).Value
    metric_sheet.Cells((last_entry + 1), 2).Value = metric_sheet.Cells(last_entry, 2).Value
    
    # Copying actual data into the cells 
    metric_sheet.Cells((last_entry + 1), 1).Value = date
    metric_sheet.Cells((last_entry + 1), 2).Value = f"='{date}'!01"

    print("Metric update complete")
    # Save and close workbook
    time.sleep(_QUIT_DELAY)

    daily_wb.Save()
    daily_wb.Close()
    print("Daily saved")

    # Close the workbook
    weekly_wb.Save()
    weekly_wb.Close()
    print("Weekly saved")
    
    # Quit Excel
    excel.Quit()

    print("Updated")

def saving(source, third_month, month, year, type):
    
    # Connect to Excel using the win32com.client package
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    op_workbook = excel.Workbooks.Open(source)
    data_workbook = excel.Workbooks.Add()

    excel.Visible=True
    excel.DisplayAlerts=False

    # Loop through the sheets in the source workbook
    for sheet in op_workbook.Sheets:
        if third_month in sheet.Name:
            sheet.Copy(None, data_workbook.Sheets(data_workbook.Sheets.Count))

    # Loop through the sheets in the source workbook
    for sheet in op_workbook.Sheets:
        if third_month in sheet.Name:
            sheet.Delete()

    time.sleep(_DELAY)
    op_workbook.Close(SaveChanges=True)
    time.sleep(_QUIT_DELAY)

    time.sleep(_DELAY)
    data_workbook.SaveAs(rf"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Sensor_{type}_Uptime_Report_{month}_{year}.xlsm", FileFormat=52)

    # Quit Excel
    excel.Quit()