import win32com.client
instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos") #연결상태 확인
print(instCpCybos.IsConnect)
