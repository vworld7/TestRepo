Function GetUserInput( myPrompt )
 ' This function uses Internet Explorer to
 ' create a dialog and prompt for user input.
 '
 ' Version:             2.11
 ' Last modified:       2013-11-07
 '
 ' Argument:   [string] prompt text, e.g. "Please enter your name:"
 ' Returns:    [string] the user input typed in the dialog screen
 '
 ' Written by Rob van der Woude
 ' http://www.robvanderwoude.com
 ' Error handling code written by Denis St-Pierre
     Dim objIE

     ' Create an IE object
     Set objIE = CreateObject( "InternetExplorer.Application" )

     ' Specify some of the IE window's settings
     objIE.Navigate "about:blank"
     objIE.Document.title = "Enter the Process Details" & String( 3, "." )
     objIE.ToolBar        = False
     objIE.Resizable      = False
     objIE.StatusBar      = False
     objIE.Width          = 420
     objIE.Height         = 240

     ' Center the dialog window on the screen
     With objIE.Document.parentWindow.screen
         objIE.Left = (.availWidth  - objIE.Width ) \ 2
         objIE.Top  = (.availHeight - objIE.Height) \ 2
     End With

     ' Wait till IE is ready
     Do While objIE.Busy
         WScript.Sleep 200
     Loop
     ' Insert the HTML code to prompt for user input
     objIE.Document.body.innerHTML = "<div align=""center""><p>" & myPrompt _
                                   & "</p>" & vbCrLf _
                                   & "<p><input type=""text"" size=""50"" " _
                                   & "id=""UserInput""></p>" & vbCrLf _
                                   & "<p><input type=""hidden"" id=""OK"" " _
                                   & "name=""OK"" value=""0"">" _
                                   & "<input type=""submit"" value="" OK "" " _
                                   & "OnClick=""VBScript:OK.value=1""></p></div>"
     ' Hide the scrollbars
     objIE.Document.body.style.overflow = "auto"
     ' Make the window visible
     objIE.Visible = True
     ' Set focus on input field
     objIE.Document.all.UserInput.focus

     ' Wait till the OK button has been clicked
     On Error Resume Next
     Do While objIE.Document.all.OK.value = 0 
         WScript.Sleep 200
         ' Error handling code by Denis St-Pierre
         If Err Then ' user clicked red X (or alt-F4) to close IE window
             IELogin = Array( "", "" )
             objIE.Quit
             Set objIE = Nothing
             Exit Function
         End if
     Loop
     On Error Goto 0


     ' Read the user input from the dialog window
     GetUserInput = objIE.Document.all.UserInput.value

     ' Close and release the object
     objIE.Quit
     Set objIE = Nothing
 End Function


sharedPath = GetUserInput("<br/><h3>Enter Shared Drive Path:</h3>" )

OutputValue=sharedPath

'WScript.Echo "Shared Path is: " & sharedPath

WScript.StdOut.WriteLine(OutputValue)
