#Author Lesley Chingwena
#Description: This script copies data from uptime csv file to the KIP report "raw_data"
import os
import win32com.client as win32
import csv
import time
import re
import gc

_DELAY = 100.0  # second
_PRO_DELAY = 3.0
_QUIT_DELAY = 10.0

def update(source, weekly, daily, health, sishen, metric, date):

    # Load the source workbook
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts=False
    
    source_wb = excel.Workbooks.Add()
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

    # Get the existing sheet "Metric" in the existing workbook
    raw_data_sheet = weekly_wb.Sheets("data_raw")

    # Clear the existing data in the "Metric" sheet
    raw_data_sheet.Cells.ClearContents()

    # Copy the data from the source sheet to the raw data sheet
    raw_rows = source_sheet.UsedRange.Rows.Count
    source_sheet.Range(f"A1:Q{raw_rows}").Copy(raw_data_sheet.Range("A1"))
    source_wb.Close(SaveChanges=False)

    # Run the macro that corresponds to the VBA button (Copy New Uptime Data)
    excel.Run('Sheet1.CommandButton1_Click')

    # Delay to allow for KPI calculations before further manupulation
    time.sleep(_DELAY)
    daily_wb = excel.Workbooks.Open(daily, ReadOnly=False)
    
    # Update Daily KPI
    daily_uptime_update(weekly_wb, daily_wb, date)
    
    # Close the workbook
    time.sleep(_QUIT_DELAY)
    weekly_wb.Save()
    weekly_wb.Close()
    print("Weekly saved")

    # Sishen uptime report update
    sishen_wb = excel.Workbooks.Open(sishen, ReadOnly=False)
    sishen_update(date, daily_wb, sishen_wb)
    time.sleep(_PRO_DELAY)
    print("Sishen uptime report complete")

    # Update health report
    health_wb = excel.Workbooks.Open(health, ReadOnly=False)
    health_update(daily_wb, health_wb, date)
    time.sleep(_PRO_DELAY)
    print("health workbook update complete")

    # Update daily metric
    metric_wb = excel.Workbooks.Open(metric, ReadOnly=False)
    daily_metric_update(daily_wb, date, metric_wb)
    print("Daily Metric update complete")

    # Save and close workbook
    time.sleep(_QUIT_DELAY)
    daily_wb.Save()
    daily_wb.Close()
    print("Daily saved")

    # Quit Excel
    excel.Quit()
    # Release Excel objects and force garbage collection
    del dm_workbook, source_wb
    gc.collect()

    print("Updated")

def sishen_update(date, tdy_wb, sis_wb):
    tdy_ws = tdy_wb.Sheets(f"{date}")

    # Add new worksheet to Sishen workbook 
    sis_ws = sis_wb.Sheets.Add(After=sis_wb.Sheets(sis_wb.Sheets.Count))
    sis_ws.Name = f'{date}'

    # Identify the sishen sensors 
    row_range = tdy_ws.UsedRange.Rows.Count
    key = "kio_sis"
    sis_rows = []

    # identifier
    for x in range(1, row_range+1):
        cell_value = tdy_ws.Cells(x, 1).Value
        if cell_value and key in  cell_value.lower():
            sis_rows.append(x)

    # Copy headings and cell formatting
    sis_wb.Sheets("Format").Range("A1:T3").Copy(sis_ws.Range("A1"))
    sis_wb.Sheets("Format").Range("S4:T14").Copy(sis_ws.Range("S4"))

    # Copy data to Sishen uptime report
    tdy_ws.Range(f"A{min(sis_rows)}:R{max(sis_rows)}").Copy(sis_ws.Range("A4"))

    time.sleep(_QUIT_DELAY)
    sis_wb.Save()
    sis_wb.Close()

def health_update(src_wb, dst_wb, date):
    # Copy daily uptime to health report
    src_ws = src_wb.Sheets(f"{date}")
    src_ws.Copy(Before=dst_wb.Sheets("Result"))
    sheet_new = dst_wb.Sheets(f"{date}")

    # Delete the "new" sheet if it exists
    
    dst_wb.Sheets("Result").Delete()
    # rename the new copied worksheet
    sheet_new.Name = "Result"

    dst_wb.Save()
    dst_wb.Close()

def daily_metric_update(src_wb, date, metric_wb):
    
    metric_ws = metric_wb.Sheets("7 days")
    src_ws = src_wb.Sheets("Metric")

    # Get max row from Matrix worksheet
    max_row = src_ws.UsedRange.Rows.Count
    kpi = src_ws.Cells(max_row, 2)

    # Get the number of used rows in the first three columns
    last_row = metric_ws.UsedRange.Rows.Count

    # Loop through the used rows
    prev_entry = metric_ws.Range(f"A{last_row}", f"C{last_row}").Copy()
    metric_ws.Range(f"A{last_row + 1}", f"C{last_row + 1}").PasteSpecial(-4122)

    # Update matrix data
    metric_ws.Cells(last_row + 1, 2).Value = kpi
    metric_ws.Cells(last_row + 1, 1).Value = date
    metric_ws.Cells(3, 3).Value = ''
    metric_ws.Cells(9, 3).Value = '=SUM(B9-B8)'

    metric_ws.Rows(2).EntireRow.Delete()

    time.sleep(_QUIT_DELAY)
    metric_wb.Save()
    metric_wb.Close()

def daily_uptime_update(weekly_wb, daily_wb, date):
    
    tdy_sheet = weekly_wb.Sheets(f"{date}")

    # Determine the size of data in the existing sheet
    data_range = tdy_sheet.UsedRange

    # Copy data from existing sheet to new sheet 
    tdy_sheet.Copy(Before=daily_wb.Sheets("VBA"))

    new_sheet = daily_wb.Sheets(f"{date}")

    # Copy formatting from previous sheet 
    format_sheet = daily_wb.Sheets("Format")
    s_range = data_range.Columns("S").Rows.Count
    format_sheet.Range(f"S1:S{s_range}").Copy(new_sheet.Range("S1"))

    # Update Metrics sheet 
    metric_sheet = daily_wb.Sheets("Metric")
    row_max = metric_sheet.UsedRange.Rows.Count
    
    # Copying formatting into the folowing row
    prev_entry = metric_sheet.Range(f"A{row_max}", f"B{row_max}").Copy()
    metric_sheet.Range(f"A{row_max + 1}", f"B{row_max + 1}").PasteSpecial(-4163)
    
    # Copying actual data into the cells 
    metric_sheet.Cells((row_max + 1), 1).Value = date
    cell_val = f"='{date}'!$O$1"
    metric_sheet.Cells((row_max + 1), 2).Value = cell_val

def saving(source, third_month, year, type):
    
    # Connect to Excel using the win32com.client package
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    op_workbook = excel.Workbooks.Open(source)
    data_workbook = excel.Workbooks.Add()

    excel.Visible=False
    excel.DisplayAlerts=False

    # Loop through the sheets in the source workbook and copy to backup
    sheet_names = [sheet.Name for sheet in op_workbook.Sheets if re.search(third_month, sheet.Name)]
    for sheet_name in sheet_names:
         sheet = op_workbook.Sheets(sheet_name)
         sheet.Copy(None, data_workbook.Sheets(data_workbook.Sheets.Count))
         sheet.Delete()

    time.sleep(_QUIT_DELAY)
    op_workbook.Save()
    op_workbook.Close()
    data_workbook.SaveAs(rf"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Monthly\Sensor_{type}_Uptime_Report_{third_month}_{year}.xlsm", FileFormat=52)

    # Quit Excel
    excel.Quit()

