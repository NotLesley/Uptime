import xlwings as xw
import win32com.client as win32c

# Assign the name of the data sheet to a variable
data_sheet = "data_raw"

# Use xlwings to get the name of the new sheet with the current date formatted
new_sheet = xw.Range("A1").parent.parent.api.Evaluate(f"TEXT(TODAY(),\"d_MMMM_YYYY\")")

# Initialize xlwings App with visibility and book creation set to False
app = xw.App(visible=True, add_book=False)

# Load the workbook
workbook = xw.Book("path_to_your_workbook.xlsx")

# Try to delete the sheet with the same name as `new_sheet`
# If it doesn't exist, the error is ignored and the code continues
try:
    workbook.sheets[new_sheet].delete()
except:
    pass

# Add a new sheet with the name `new_sheet` to the workbook
workbook.sheets.add(new_sheet)

----------------------------------------------------------------------------

# Define the source and target worksheets
WshSrc = workbook.sheets[prev_sheet_name]
WshTrg = workbook.sheets[new_sheet]

# Copy the used range from the source worksheet
WshSrc.api.UsedRange.Copy()

# Paste the column widths, formats, and formulas/number formats
# into the target worksheet
WshTrg.api.Range("A1").PasteSpecial(Paste=win32c.xlPasteColumnWidths)
WshTrg.api.Range("A1").PasteSpecial(Paste=win32c.xlPasteFormats)
WshTrg.api.Range("A1").PasteSpecial(Paste=win32c.xlPasteFormulasAndNumberFormats)

------------------------------------------------------------------------------

# Constants
data_sheet = "data_raw"

# Get workbook and worksheets
app = xlwings.App(visible=True, add_book=False)
workbook = xw.Book("path_to_your_workbook.xlsx")
WshTrg = workbook.sheets[new_sheet]
WshData = workbook.sheets[data_sheet]

# Get target and data ranges
target_range = WshTrg.range("A5:A700")
data_range = WshData.range("A2:A700")

# Loop through target range
for cell in target_range:
    if len(cell.value) > 6:
        match_found = 0
        value_to_find = cell.value
        
        # Loop through data range
        for data_cell in data_range:
            if data_cell.value == value_to_find:
                match_found = 1
                
                # Check if wsps or server is in the value
                pos_wsps = value_to_find.find("wsps")
                pos_server = value_to_find.find("server")
                
                # Copy data from data sheet to new sheet
                WshData.range(f"B{data_cell.row}:I{data_cell.row}").api.Copy(WshTrg.range(f"B{cell.row}:I{cell.row}"))
                WshTrg.range(f"B{cell.row}:I{cell.row}").number_format = "0.0"
                
                # If wsps or server is not in the value, copy additional data
                if pos_wsps + pos_server == 0:
                    WshData.range(f"N{data_cell.row}:Q{data_cell.row}").api.Copy(WshTrg.range(f"J{cell.row}:M{cell.row}"))
                    WshTrg.range(f"J{cell.row}:M{cell.row}").number_format = "0.0"
                    
        # If no match is found, show a message
        if match_found == 0:
            xw.apps[app].api.MsgBox(f"No match found for {cell.value}")
            
# Loop through used range in new sheet and change number format if cell is numeric
for r in WshTrg.used_range.special_cells(xlCellTypeConstants):
    if isinstance(r.value, float):
        r.value = float(r.value)
        r.number_format = "0.0"

-----------------------------------------------------------------------------------------

# Define the metric sheet
metric_sheet = "Metric"

# Get the last row of the metric sheet
lastrow = xw.Range(metric_sheet, 'A').end('up').row + 1

# Write the new_sheet name in column A of the last row
xw.Range(metric_sheet, lastrow, 1).value = new_sheet

# Write the value of cell O1 of the new_sheet in column B of the last row
xw.Range(metric_sheet, lastrow, 2).value = "='" + new_sheet + "'!O1"

# Pop up a message to show that the process has finished
xw.msgbox("Finished")

# Copy the value of cell C2 from the previous sheet to cell B1 of the new_sheet
prev_sheet_name = xw.sheets[new_sheet].previous.name
prev_sheet = xw.sheets[prev_sheet_name]
prev_sheet.range("C2").value = xw.sheets[new_sheet].range("B1").value