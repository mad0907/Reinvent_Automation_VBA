import win32com.client
xl = win32com.client.Dispatch("Excel.Application")  #instantiate excel app
vMacroPath="C:/Users/munmun.a.das/OneDrive - Accenture/Documents/Python Scripts/TestMacro.xlsm"
vStrArg="{Name:Munmun, Marks:90}"
wb = xl.Workbooks.Open(vMacroPath)
xl.Application.Run('TestMacro.xlsm!Module1.RunFromPython1("'+vStrArg+'")')
wb.Save()
xl.Application.Quit()
