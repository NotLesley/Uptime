import win32com.client as win32
from datetime import date
from download import Down

# Date
today = date.today()

dm_source = r""

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts=False

source_wb = excel.Workbooks.Add()
source_sheet = source_wb.Sheets("Metric")

dm_workbook = excel.Workbooks.Open(dm_source, ReadOnly=False)
dm_sheet = dm_workbook.Sheets("recent")

# Get max row from Matrix worksheet
max_row = dm_sheet.Cells(dm_sheet.Rows.Count, 1).End(-4162).Row
kpi = source_sheet.Cells(max_row, 2)

# Get the number of used rows in the first three columns
last_row = dm_sheet.Cells(dm_sheet.Rows.Count, 1).End(-4162).Row

# Loop through the used rows
for i in range(2, last_row + 1):
    # Copy the values from the current row
    values = [dm_sheet.Cells(i, j).Value for j in range(1, 4)]

    # Paste the values into the next row
    for j, value in enumerate(values, start=1):
        dm_sheet.Cells(i + 1, j).Value = value

# Update matrix data
dm_sheet.Cells(9, 1).Value = today
dm_sheet.Cells(9, 3).Value = kpi

for x in range(3, 9):
    y = x - 1
    z = x + 1
    cell_val = f'=SUM(B{x}-B{y})'
    dm_sheet.Cells(3, 3).Value = ''
    dm_sheet.Cell(z, 3).Value = cell_val

dm_sheet.Rows(2).EntireRow.Delete()

dm_workbook.Save()
dm_workbook.Close()

print("Daily matrix updated")