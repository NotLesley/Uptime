import win32com.client as win

xl = win.DispatchEx("Excel.Application")

source = r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Book1.xlsm"

wb = xl.Workbooks.Open(source)

xl.Visible = True

xl.Run("Sheet1.CommandButton1_Click")

wb.SaveAs(r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Book2.xlsm")

wb.Close()

xl.Quit()
