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
HttpErr = "<p><img border=�0� src=�res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/pagerror.gif�>&nbsp;"
HttpErr = HttpErr + "<span id=�errorText�><font size=�4� color=�#FF0000�>The Page You Selected cannot"
HttpErr = HttpErr + " be Displayed</font></span></p>"
HttpErr = HttpErr + "<p><b>Peronal Web Server</b></p>"
HttpErr = HttpErr + "<p>The page you are looking for is currently unavailable. The Web&nbsp;<br>"
HttpErr = HttpErr + "site might be experiencing technical difficulties, or you may need&nbsp;<br>"
HttpErr = HttpErr + "to adjust your browser settings.</p>"
HttpErr = HttpErr + "<hr>"
HttpErr = HttpErr + "<p>Try one of the flowing options below:</p>"
HttpErr = HttpErr + "<ul>"
HttpErr = HttpErr + " <li><font size=�3�>Please make sure that the address you type in the URL bar"
HttpErr = HttpErr + " is spelled correctly.</font></li>"
HttpErr = HttpErr + "  <li><font size=�3�>Click <a href=�javascript:history.back(1)�><img border=�0� src=�res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/back.gif�>"
HttpErr = HttpErr + "    Back</a> and try again latter</font></li>"
HttpErr = HttpErr + "  <li><font size=�3�>Click here <a href=�javascript:location.reload()�><img border=�0� src=�res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/refresh.gif�></a>"
HttpErr = HttpErr + "    <a href=�javascript:location.reload()�>refresh</a> and try again</font></li>"
HttpErr = HttpErr + "</ul>"
HttpErr = HttpErr + "<hr>"
HttpErr = HttpErr + "<p><i>If you have any more errors then please<br>"
HttpErr = HttpErr + "contact the local admin service</i></p>"
HttpErr = HttpErr + "<p><u><b>HTTP 404 - Page Not Found</b></u></p>"

Http_404_Error = Replace(HttpErr, Chr(255), Chr(34))

End Sub


