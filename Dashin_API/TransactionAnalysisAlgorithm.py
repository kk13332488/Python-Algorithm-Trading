import win32com.client
import time
def CheckVolume(instStockChart, code):
    instStockChart.SetInputValue(0, code)
    instStockChart.SetInputValue(1, ord('2'))
    instStockChart.SetInputValue(4, 60) #last 60 days transaction amount data
    instStockChart.SetInputValue(5, 8)
    instStockChart.SetInputValue(6, ord('D'))
    instStockChart.SetInputValue(9, ord('1'))

    instStockChart.BlockRequest()

    volumes = []
    numData = instStockChart.GetHeaderValue(3)
    for i in range(numData):
        volume = instStockChart.GetDataValue(0, i)
        volumes.append(volume)

    averageVolume = (sum(volumes) - volumes[0]) / (len(volumes) -1) #59 days transaction amount / 59

    if (volumes[0] > averageVolume * 10):
        return 1
    else:
        return 0

if __name__ == "__main__":
    instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = instCpCodeMgr.GetStockListByMarket(1)
    buyList = []
    for code in codeList:
        if CheckVolume(instStockChart, code) == 1:
            buyList.append(code)
            print(code)
        time.sleep(1)



