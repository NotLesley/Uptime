import win32com.client as win32
import datetime
import time

_QUIT_DELAY = 10
weekly_uptime = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Weekly_Uptime_Report.xlsm"
sis_uptime = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Sishen_Uptime.xlsm"

def add_data(tdy_sheet):
    # Identify the sishen sensors 
    row_range = tdy_sheet.UsedRange.Rows.Count
    key = "kio_sis"
    sis_rows = []

    # identifier
    for x in range(1, row_range+1):
        cell_value = tdy_sheet.Cells(x, 1).Value
        if cell_value and key in  cell_value.lower():
            sis_rows.append(x)

    # Copy headings 
    sis_wb.Sheets("Format").Range("A1:T3").Copy(sis_ws.Range("A1"))
    sis_wb.Sheets("Format").Range("S4:T14").Copy(sis_ws.Range("S4"))

    # Copy data to Sishen uptime report
    tdy_sheet.Range(f"A{min(sis_rows)}:R{max(sis_rows)}").Copy(sis_ws.Range("A4"))

if __name__ == '__main__':

    excel = win32.gencache.EnsureDispatch('Excel.Application')

    weekly_wb = excel.Workbooks.Open(weekly_uptime, ReadOnly=False)
    sis_wb = excel.Workbooks.Open(sis_uptime, ReadOnly=False)

    now = datetime.datetime.now()
    date_string = now.strftime("%d %b %Y")

    # Add new worksheet to Sishen workbook 
    sis_ws = []
    for sheet in weekly_wb.Sheets:
        if "Feb" in sheet.Name:
            sis_ws.append(sis_wb.Sheets.Add(After=sis_wb.Sheets(sis_wb.Sheets.Count)))
            sis_ws[len(sis_ws) - 1].Name = sheet.Name
            add_data(sheet)
    # for x in range(len(sis_ws)):
    #     sis_ws[x].Name = sheet.Name


    time.sleep(_QUIT_DELAY)

    weekly_wb.Close()
    sis_wb.Save()
    sis_wb.Close()

    print("ALL OKAY")