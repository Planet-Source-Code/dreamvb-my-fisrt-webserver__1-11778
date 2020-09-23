Attribute VB_Name = "Module1"
Public Http_404_Error As String

Public Function FileExists(ByVal sFilename As String) As Integer
 Dim x
  x = Dir(sFilename)
  
  If x = "" Then
    FileExists = 0
    Else
    FileExists = -1
  End If
  
End Function
Sub Http404()
Dim HttpErr As String
HttpErr = ""
HttpErr = "<p><img border=ÿ0ÿ src=ÿres://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/pagerror.gifÿ>&nbsp;"
HttpErr = HttpErr + "<span id=ÿerrorTextÿ><font size=ÿ4ÿ color=ÿ#FF0000ÿ>The Page You Selected cannot"
HttpErr = HttpErr + " be Displayed</font></span></p>"
HttpErr = HttpErr + "<p><b>Peronal Web Server</b></p>"
HttpErr = HttpErr + "<p>The page you are looking for is currently unavailable. The Web&nbsp;<br>"
HttpErr = HttpErr + "site might be experiencing technical difficulties, or you may need&nbsp;<br>"
HttpErr = HttpErr + "to adjust your browser settings.</p>"
HttpErr = HttpErr + "<hr>"
HttpErr = HttpErr + "<p>Try one of the flowing options below:</p>"
HttpErr = HttpErr + "<ul>"
HttpErr = HttpErr + " <li><font size=ÿ3ÿ>Please make sure that the address you type in the URL bar"
HttpErr = HttpErr + " is spelled correctly.</font></li>"
HttpErr = HttpErr + "  <li><font size=ÿ3ÿ>Click <a href=ÿjavascript:history.back(1)ÿ><img border=ÿ0ÿ src=ÿres://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/back.gifÿ>"
HttpErr = HttpErr + "    Back</a> and try again latter</font></li>"
HttpErr = HttpErr + "  <li><font size=ÿ3ÿ>Click here <a href=ÿjavascript:location.reload()ÿ><img border=ÿ0ÿ src=ÿres://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/refresh.gifÿ></a>"
HttpErr = HttpErr + "    <a href=ÿjavascript:location.reload()ÿ>refresh</a> and try again</font></li>"
HttpErr = HttpErr + "</ul>"
HttpErr = HttpErr + "<hr>"
HttpErr = HttpErr + "<p><i>If you have any more errors then please<br>"
HttpErr = HttpErr + "contact the local admin service</i></p>"
HttpErr = HttpErr + "<p><u><b>HTTP 404 - Page Not Found</b></u></p>"

Http_404_Error = Replace(HttpErr, Chr(255), Chr(34))

End Sub


