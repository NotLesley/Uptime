import win32com.client as win32

daily_uptime = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Daily_Uptime_Report.xlsm"
weekly_uptime =  r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Weekly_Uptime_Report.xlsm"

excel = win32.gencache.EnsureDispatch('Excel.Application')

weekly_wb = excel.Workbooks.Open(weekly_uptime, ReadOnly=False)
daily_wb = excel.Workbooks.Open(daily_uptime, ReadOnly=False)
excel.Visible = True
excel.DisplayAlerts = False

today_sheet = weekly_wb.Sheets("13 Feb 2023")

today_sheet.Copy(Before=daily_wb.Sheets("VBA"))

new_sheet = daily_wb.Sheets("13 Feb 2023")

data_range = today_sheet.UsedRange

prev_sheet = daily_wb.Sheets("10 Feb 2023")
s_range = data_range.Columns("S").Rows.Count
prev_sheet.Range(f"S1:S{s_range}").Copy(new_sheet.Range("S1"))
print("formatting complete")

daily_wb.Save()
daily_wb.Close()

weekly_wb.Save()
weekly_wb.Close()

excel.Quit()

print("ALL OKAY")