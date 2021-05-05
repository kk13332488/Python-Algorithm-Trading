import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open("C:\\input.xlsx") #openning the local excel file
ws = wb.ActiveSheet
print(ws.Cells(1,1).Value) #print local excel file's contents
excel.Quit()
