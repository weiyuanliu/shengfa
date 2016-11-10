<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　　七日科技企业网站管理系统（LISuo）　　　　　　　  ┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
' 　版权所有　qisehu.com
'
'　　程序制作　七日科技有限公司
'　　　　　　　Add:四川省成都市二环路西三段181号13楼20/21号
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%>
<% Option Explicit %>
<%
Dim Action : Action=Request.QueryString("Action")
Select Case Action
	Case "ver","Ver"
		Main() : ShowVer()
	Case Else
		Main()
End Select

Sub Main()
%>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2007-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="七日科技企业网站,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>管理员登录</TITLE>
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>

<BODY style="margin-top:100px;font-size:12px;">
<div align="center"><img src="Images/Login_Top.jpg" width="530" height="150"></div>
<div align="center">
  <table width="530" height="100" border="0" cellpadding="0" cellspacing="0" background="Images/Login_Bottom.jpg">
	<form action="CheckLogin.asp" method="post" name="AdminLogin" id="AdminLogin"  onSubmit="return CheckAdminLogin()">
    <tr>
      <td width="70" height="46" rowspan="2">&nbsp;</td>
      <td width="132" rowspan="2" valign="bottom">
      <input name="LoginName" type="text" id="LoginName" maxlength="12" style="width:94px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7" onMouseOver="this.style.background='#ffffff'" onMouseOut="this.style.background='#F7F7F7'" onFocus="this.select();"></td>
      <td width="131" rowspan="2" valign="bottom">
      <input name="LoginPassword" type="password" id="LoginPassword" maxlength="12" style="width:94px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7" onMouseOver="this.style.background='#ffffff'" onMouseOut="this.style.background='#F7F7F7'" onFocus="this.select();"></td>
      <td width="60" rowspan="2" valign="bottom">
      <input name="VerifyCode" type="text" id="BC27F457E1FA34EF93" maxlength="4" style="width:60px; BORDER-RIGHT: #F7F7F7 0px solid; BORDER-TOP: #F7F7F7 0px solid; FONT-SIZE: 9pt; BORDER-LEFT: #F7F7F7 0px solid; BORDER-BOTTOM: #c0c0c0 1px solid; HEIGHT: 16px; BACKGROUND-COLOR: #F7F7F7" onMouseOver="this.style.background='#ffffff'" onMouseOut="this.style.background='#F7F7F7'" onFocus="this.select();">	  </td>
      <td width="62" height="25" valign="bottom"><img src="../Include/VerifyCode.asp" align="absmiddle"></td>
      <td width="75" valign="bottom">
	  <input name="submitLogin" type="image" src="images/Login_Submit.jpg" width="40" height="34"></td>
    </tr>
	</form>
    <tr>
      <td height="1" valign="bottom"></td>
      <td width="75" valign="bottom"></td>
    </tr>
    <tr>
      <td height="54" colspan="6">&nbsp;</td>
    </tr>
  </table>
<%
End Sub

Sub ShowVer()
	Response.Write("<div>"& vbcrlf)
	Response.Write("<script>"& vbcrlf)
	Response.Write("<!--"& vbcrlf)
	Response.Write("document.write(unescape('%3Cscript%3E%0D%0A%3C%21--%0D%0Adocument.write%28unescape%28%22%25u5236%25u4F5C%253A%25u51AF%25u70B3%25u57FA%22%29%29%3B%0D%0A//--%3E%0D%0A%3C/script%3E'));"& vbcrlf)
	Response.Write("//-->"& vbcrlf)
	Response.Write("</script>"& vbcrlf)
	Response.Write("</div>")
End Sub
%>
</div>
</BODY>
</HTML>
