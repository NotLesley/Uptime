import xlwings as xw

app = xw.App(visible=True, add_book=False)
workbook = xw.Book(r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Sensor_Uptime_Report_02_Feb_2023.xlsm")

xw.books[r"C:\Users\Lesley Chingwena\Documents\python_scripts\Uptime\docs\Sensor_Uptime_Report_02_Feb_2023.xlsm"].app.run("VBA.CommandButton1_Click")

wb.save()
wb.app.quit()
