#Author Lesley Chingwena
#Description: This script copies data from uptime csv file to the KIP report "raw_data"
import os
import win32com.client as win32
import csv
import time

_DELAY = 100.0  # second
_QUIT_DELAY = 10.0

def update(source, target):

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
    target_wb = excel.Workbooks.Open(target, ReadOnly=False)
    excel.Visible = False
    excel.DisplayAlerts=False

    # Get the existing sheet "Metric" in the existing workbook
    raw_data_sheet = target_wb.Sheets("data_raw")

    # Clear the existing data in the "Metric" sheet
    raw_data_sheet.Cells.ClearContents()

    # Copy the data from the source sheet to the "Metric" sheet
    source_sheet.Range("A1:Q735").Copy(raw_data_sheet.Range("A1"))

    # Close the workbooks
    source_wb.Close(SaveChanges=False)

    # Run the macro that corresponds to the VBA button (Copy New Uptime Data)
    excel.Run('Sheet1.CommandButton1_Click')

    # Delay to allow for KPI calculations before saving and closing excel
    time.sleep(_DELAY)


    # if os.path.exists(target):
    #     os.remove(target)

    # Close the workbook
    target_wb.Save()
    target_wb.Close()
    
    # Create an instance of the WScript.Shell object
    # shell = win32.Dispatch("WScript.Shell")

    # # Use the SendKeys method to simulate keystrokes
    # # The TAB key is used to move the focus to the "Yes" button in the prompt
    # shell.SendKeys("{TAB}")

    # # The ENTER key is used to activate the "Yes" button and replace the existing file
    # shell.SendKeys("{ENTER}")
        
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

    # Create an instance of the WScript.Shell object
    shell = win32.Dispatch("WScript.Shell")

    # Use the SendKeys method to simulate keystrokes
    # The TAB key is used to move the focus to the "Yes" button in the prompt
    shell.SendKeys("{TAB}")

    # The ENTER key is used to activate the "Yes" button and replace the existing file
    shell.SendKeys("{ENTER}")

    time.sleep(_DELAY)
    data_workbook.SaveAs(rf"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Sensor_{type}_Uptime_Report_{month}_{year}.xlsm", FileFormat=52)

    # Quit Excel
    excel.Quit()

