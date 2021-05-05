import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Add() #Creating worksheet
ws = wb.Worksheets("Sheet1")
ws.Cells(1,1).Value = "Hello World!" #
wb.SaveAs("c:\\Users\\고성호\\Desktop\\test.xlsx") #Saving and setting the path
excel.Quit()