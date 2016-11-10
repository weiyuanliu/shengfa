<%
Sub Check_Wap()
 dim MoblieUrl,reExp,MbStr
 MoblieUrl = "/wap/"
 Set reExp = New RegExp
 MbStr = "Android|iPhone|UC|Windows Phone|webOS|BlackBerry|iPod"
 reExp.pattern = ".*("&MbStr&").*"
 reExp.IgnoreCase = True
 reExp.Global = True
 If reExp.test(Request.ServerVariables("HTTP_USER_AGENT")) Then
  response.redirect MoblieUrl
  response.End
 End If
End Sub
%>