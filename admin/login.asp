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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0055)http://leven.demo.salesproduct.cn/backoffice/home.phtml -->
<HTML><HEAD><TITLE>顺意网络后台管理系统</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="login.files/touch-maker-style1.css" type=text/css rel=stylesheet>
<script language="javascript" src="../Script/Admin.js"></script>

<META content="MSHTML 6.00.2800.1589" name=GENERATOR><style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
.STYLE4 {
	font-size: 16px;
	color: #CC3300;
	font-weight: bold;
}
-->
</style></HEAD>
<BODY leftMargin=0 topMargin=0  MARGINHEIGHT="0" 
MARGINWIDTH="0">
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD align=right background=login.files/touch-maker_5.gif>
      <DIV class=pc_left></DIV></TD>
    <TD width=868 height=531>
      <TABLE height=531 cellSpacing=0 cellPadding=0 width=868 border=0>
        <TBODY>
        <TR>
          <TD width=291 background=login.files/touch-maker_10.gif>&nbsp;</TD>
          <TD class=wai vAlign=bottom width=577 
          background=login.files/touch-maker_1.gif>
            <DIV class=demo></DIV>
            <TABLE height=200 cellSpacing=0 cellPadding=0 width=450 border=0>
              <TBODY>
              <TR>
                <TD width=165 height=176>
                  <TABLE height=171 cellSpacing=0 cellPadding=0 width=165 
                  border=0>
                    <TBODY>
                    <TR>
                      <TD height=94 align="left" vAlign=bottom><span class="STYLE4">顺意网络<br>
                        <br>
                        后台管理系统</span></TD>
                    </TR>
                    <TR>
                      <TD vAlign=center align=left>&nbsp;</TD>
                    </TR></TBODY></TABLE></TD>
                <TD vAlign=top width=285>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
				  <form action="CheckLogin.asp?t=<%=request("t")%>" method="post" name="AdminLogin" id="AdminLogin"  onSubmit="return CheckAdminLogin()">
                    <TBODY>
                    <TR>
                      <TD vAlign=top align=right width="29%" 
                      height=24>&nbsp;</TD>
                      <TD vAlign=top width="71%"><SPAN class=red></SPAN></TD></TR>
                    <TR>
                      <TD align=right height=35><IMG height=31 
                        alt="Identity / 帐 号 :" 
                        src="login.files/touch-maker_icon.gif" width=33></TD>
                      <TD class=right align=left><INPUT name=LoginName class=or id="LoginName"> 
                      </TD></TR>
                    <TR>
                      <TD align=right height=35><IMG height=28 
                        alt="Password / 密 码 :" 
                        src="login.files/touch-maker_icon1.gif" width=30></TD>
                      <TD class=right><INPUT 
                        name=LoginPassword type=password class=or id="LoginPassword"> </TD></TR>
						  <TR>
                      <TD align=right height=35>验证：</TD>
                      <TD class=right><INPUT 
                        name="VerifyCode" type=text style="width:80px" class=or id="LoginPassword">&nbsp;&nbsp;<img src="../Include/VerifyCode.asp" alt="看不清楚?请点击刷新" onclick="this.src=this.src+'?'+Math.random();" align="absmiddle" style="CURSOR:hand;"></TD></TR>
						






                    <TR>
                      <TD align=right>&nbsp;</TD>
                      <TD class=right1 align=middle><INPUT 
                       type=image alt=登入 
                        src="login.files/touch-maker_icon5.gif" border=0> 
                    </TD></TR></TBODY>	</form></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
    <TD background=login.files/touch-maker_7.gif>
      <DIV class=pc_right></DIV></TD></TR>
  <TR>
    <TD background=login.files/touch-maker_8.gif>&nbsp;</TD>
    <TD vAlign=top background=login.files/touch-maker_3.gif>
      <DIV class=pc_bottom>
      <TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
        <TBODY>
        <TR>
          <TD align=left width="46%"></TD>
          <TD class=right align=middle width="17%"></TD>
          <TD class=right align=left width="37%">&nbsp;</TD>
        </TR></TBODY></TABLE></DIV></TD>
    <TD background=login.files/touch-maker_9.gif>&nbsp;</TD></TR><INPUT 
  type=hidden value=yes name=Myself> <INPUT type=hidden name=App> 
</FORM></TBODY></TABLE><%
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
%></BODY></HTML>
