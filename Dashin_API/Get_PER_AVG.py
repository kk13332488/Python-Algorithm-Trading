import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

targetCodeList = instCpCodeMgr.GetGroupCodeList(5) #code list of section 5

instMarketEye.SetInputValue(0, 67) #inputvalue = PER
instMarketEye.SetInputValue(1, targetCodeList)

instMarketEye.BlockRequest()

numStock = instMarketEye.GetHeaderValue(2) #number of stocks

sumPer = 0
for i in range(numStock):
    sumPer += instMarketEye.GetDataValue(0, i)

print("Average PER: ", sumPer / numStock)
