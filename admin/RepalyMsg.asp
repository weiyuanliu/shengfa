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
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>审核、修改、回复留言</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
call CreateEditor("Content")
%>


<%
if Instr(session("AdminPurview"),"|90,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result,ID
Result=request.QueryString("Result")
dim Msg_Name,Msg_Time,Msg_TelPhone,Replay,ReplayTime
ID=request.QueryString("ID")
call MesEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<STRONG>留言信息：审核，修改，回复留言信息相关的内容</STRONG></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="MessageList.asp" onClick='changeAdminFlag("留言信息列表")'>查看留言信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetSite.asp#Message" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>设置是否自动审核</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="RepalyMsg.asp?Action=SaveEdit&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">留言者：</td>
        <td><input name="Msg_Name" type="text" class="textfield" id="Msg_Name" style="WIDTH: 240;" value="<%=Msg_Name%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">留言者联系电话：</td>
        <td><input name="Msg_TelPhone" type="text" class="textfield" id="Msg_TelPhone" style="WIDTH: 240;" value="<%=Msg_TelPhone%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">留言时间：</td>
        <td><input name="Msg_Time" type="text" class="textfield" id="Msg_Time" style="WIDTH: 240" value="<%=Msg_Time%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">回复时间：</td>
        <td><input name="ReplayTime" type="text" class="textfield" id="ReplayTime" style="WIDTH: 240" value="<%if ReplayTime<>"" then response.Write(ReplayTime) else response.Write(now())%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">回复内容：</td>
        <td>
        <textarea name="Content" rows="15" class="textfield" id="Content" style="WIDTH: 100%;"  >
        	<%if Replay="" or isnull(Replay) then%>
            	<%=Msg_Name%>朋友，您的货已与<font color="#ff0000"><u> x年x月x日 </u></font>发出，您的快递单号是<font color="#ff0000"><u> xxxxxxxxxx </u></font>,在您当地为您派送产品的快递公司的联系电话<font color="#ff0000"><u> xxxxxxxx </u></font>，快递公司是<font color="#ff0000"><u> xxx快递公司 </u></font>，请及时与快递公司联系，并告诉快递公司您的快递单号，让他们及时送货给您。如果还有问题请致电400-661-9668让我们工作人员帮您解决。
            <%else%>
            	<%=Replay%>
            <%end if%>
        </textarea>
       </td>
      </tr>
	  <tr>
	  <td height="40" align="right" >操作员：</td>
	  <td><%=session("UserName")%></td>
	  </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;" >
          <input name="submitSaveEdit2" type="button" class="button"  id="submitSaveEdit2" value="返回" style="WIDTH: 80;" onClick="window.location.href=document.referrer;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom"></td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub MesEdit()
  if ID=""or isnull(ID) or not(IsNumeric(ID)) then
  	response.Write("<script langauge=javascript>"&vbcrlf)
		response.Write("alert('对不起，数据出错，请返回！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>"&vbcrlf)
	response.End()
  end if
  Dim Action,Rs,Sql,Editadmin
  Editadmin=session("UserName")
  Action=Trim(Request.QueryString("Action"))
  Set rs=server.CreateObject("adodb.recordset")
  sql="select * from MsgData where id="&id
  if Action="SaveEdit" then
  	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('记录未找到，请返回！');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
	else
		rs("Replay")=Trim(Request.Form("Content"))
		rs("ReplayTime")=Trim(Request.Form("ReplayTime"))
		rs("Ediadmin")=Editadmin
		rs.update()
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('回复留言成功！');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
	end if
  else
  	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('记录未找到，请返回！');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
  	else
		Msg_Name=rs("Msg_Name")
		Msg_Time=rs("Msg_Time")
		Msg_TelPhone=rs("Msg_TelPhone")
		Replay=rs("Replay")
		ReplayTime=rs("ReplayTime")
		Editadmin=rs("Ediadmin")
		rs.close()
		set rs=Nothing
  	end if
  end if
end sub

%>