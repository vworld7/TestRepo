SET WshShell = CreateObject("WScript.Shell")
 
SET oExec=WshShell.Exec("taskkill /F /IM iexplore.exe")
SET oExec=WshShell.Exec("taskkill /F /IM EXCEL.EXE")
 
SET oExec= Nothing
 
SET WshShell =Nothing