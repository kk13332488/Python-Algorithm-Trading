import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\input.xlsx')
ws = wb.ActiveSheet

ws.Cells(1,2).Value = "is" #writing a word to excel
ws.Range("C1").Value = "good"
ws.Range("A2:c2").Interior.ColorIndex = 10 #coloring the cell
