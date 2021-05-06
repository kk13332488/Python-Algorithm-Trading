import win32com.client
instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instStockChart.SetInputValue(0, "A003530") #Stock Code
instStockChart.SetInputValue(1, ord('2')) #Request by ASCII Code
instStockChart.SetInputValue(4, 10) #n match from last transaction date
instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 8)) #field value
instStockChart.SetInputValue(6, ord('D')) #Chart classification - Day
instStockChart.SetInputValue(9, ord('1')) #Revision stock price reflected or not

instStockChart.BlockRequest() #call server
numData = instStockChart.GetHeaderValue(3) #number of received data
numField = instStockChart.GetHeaderValue(1) #number of Field
for i in range(numData):
    for j in range(numField):
        print(instStockChart.GetDataValue(j,i), end = " ") #Get the requested data
    print("")




