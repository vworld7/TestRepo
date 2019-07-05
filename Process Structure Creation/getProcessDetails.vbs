Dim exlObj,exlFileName,sheetName,processName,botDetails,countLastRow,tempValue,i

excelfilepath = wscript.arguments(0)
'excelfilepath ="\\svrin000mbp00.asia.corp.anz.com\kurshidm$\Desktop\Automation\ReusableBot\Solution DesignIIB_PTM.xlsx"
Set exlObj=CreateObject("Excel.Application")
exlObj.Visible=false
Set exlFileName=exlObj.Workbooks.Open(excelfilepath,false)

Set sheetName=exlObj.ActiveWorkBook.WorkSheets("Index")
countLastRow=sheetName.UsedRange.Rows.Count

'MsgBox "hello"
'MsgBox countLastRow
Dim Count

For Count=1 To countLastRow
tempValue=sheetName.Cells(Count,"A").value

	If tempValue = "Project Name " Then
	'MsgBox "success"& Count
	processName=sheetName.Cells(Count,"B").value
	
	Exit For

	End If
	
Next

'MsgBox processName

sheetName=""
countLastRow=""
Count=""
tempValue=""
Set sheetName=exlObj.ActiveWorkBook.WorkSheets("Bot Description")
countLastRow=sheetName.UsedRange.Rows.Count
'MsgBox countLastRow

For Count=1 To countLastRow
tempValue=sheetName.Cells(Count,"A").value
'MsgBox tempValue
If tempValue = "Key Tasks" Then
'msgbox "Key task found"
   For i=1 To countLastRow Step 1 
	if  sheetName.Cells(Count+i,"B").value<>"" then
	   if i=1 Then
  		BotDetails=sheetName.Cells(Count+i,"B")
	   else
		BotDetails=BotDetails&"#"&sheetName.Cells(Count+i,"B")
	   end if
	end if
   Next
	
	
	'Exit For

End If
	
Next
'MsgBox (ProcessName &"#"& BotDetails)
WScript.StdOut.WriteLine(ProcessName &"@"& BotDetails)
exlFileName.close
exlObj.quit

