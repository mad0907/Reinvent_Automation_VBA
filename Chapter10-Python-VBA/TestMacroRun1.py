import win32com.client
xl = win32com.client.Dispatch("Excel.Application")  #instantiate excel app
vMacroPath="C:/Users/munmun.a.das/OneDrive - Accenture/Documents/Python Scripts/TestMacro.xlsm"
wb = xl.Workbooks.Open(vMacroPath)
xl.Application.Run('TestMacro.xlsm!Module1.RunFromPython("Hello Munmun")')
wb.Save()
xl.Application.Quit()