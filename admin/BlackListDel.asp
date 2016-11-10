<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>加入黑名单</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|305,")=0 then 
  response.write ("<script language=javascript> alert('你不具有该管理模块的操作权限，请返回！');history.back(-1);</script>")
end if
%>
<%
if Instr(session("AdminPurview"),"|305,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ReplyContent,ReplyTime,ID,ProductName,ProductNo,Amount,Remark,display,NotSend
dim Linkman,Company,Address,ZipCode,Telephone,blacklist,Mobile,Email,AddTime,States,FuKuan,HuoDao_FuKuan,Tel
ID=request.QueryString("ID")
Function Cll()
 Dim rs,sql
 set rs=server.CreateObject("Adodb.recordset")
 sql="Select top 1 OrderSates from NwebCn_Site"
 rs.open sql,conn,1,1
 if not rs.eof then
 Cll=rs(0)
 end if
End Function
call OrderEdit() 

%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>加入黑名单</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="OrderList.asp" onClick='changeAdminFlag("订单信息列表")'>查看订单信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>网站信息设置</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="BlackListDel.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" onsubmit="return checkChinese();">
  <input name="blacklist" type="hidden" id="blacklist" value="1">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">订&nbsp;购&nbsp;者：</td>
        <td><%=Linkman%></td>
      </tr>
      <tr>
        <td height="20" align="right">拉黑原因：</td>
        <td valign="bottom">

         <span style="display:<%=display%>;" id="NotSend">
         	<textarea name="NotSend" rows="6" class="textfield" id="Remark" style="WIDTH: 76%;"><%=NotSend%></textarea><br /><font color="#FF0000">*请填写原因</font>
         </span>
       </td>
      </tr>
	  <tr>
	   <td height="40" align="right">操作员:</td>
	   <td name="deladmin"><%=session("UserName") %></td>
	  </tr>
<script language="JavaScript"> 
<!-- 
function checkChinese() 
{ var isnumandchar; 
var StrForCheck=editForm.NotSend.value; 
var c; 
c = StrForCheck.charAt(0); 
while((c ==" "||c =="　") && StrForCheck.length > 0) 
{ 
StrForCheck = StrForCheck.slice(1); 
c = StrForCheck.charAt(0); 
} 
c = StrForCheck.charAt(StrForCheck.length -1); 
while((c ==" "||c =="　") && StrForCheck.length > 0) 
{ 
StrForCheck = StrForCheck.substring(0,StrForCheck.length-1); 
c = StrForCheck.charAt(StrForCheck.length -1); 
} 
editForm.NotSend.value = StrForCheck;//如果没有值，继续 
if (StrForCheck.length==0 || StrForCheck.length<2) 
{ 
if (StrForCheck.length==0) {alert("拉黑原因不能为空！");return false;} 
} 
else 
{return true;} 
} 
//--> 
</script> 
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom"><label>
          <input type="submit" name="Modify" id="Modify" value="确定拉黑">
          <input type="button" name="Modify2" id="Modify2" value="返 回" onClick="window.history.go(-1);">
        </label></td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub OrderEdit()
  dim Action,rsCheckAdd,rs,sql,deladmin
  deladmin=session("UserName")
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_Order where ID="&ID
      rs.open sql,conn,1,3
	  rs("blacklist")=Trim(Request.Form("blacklist"))
	  rs("deladmin")=deladmin
	  if Trim(Request.Form("NotSend"))<>"" then
	  	rs("NotSendText")=trim(Request.Form("NotSend"))
	  end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('已加入黑名单！');changeAdminFlag('订单信息列表');location.replace('OrderList.asp');</script>"
  else '提取留言信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Order where ID="& ID
      rs.open sql,conn,1,1
	  Linkman=GuestInfo(rs("MemID"),rs("Linkman"),rs("Sex"))
	  ProductName=rs("ProductName")
	  NotSend=rs("NotSendText")
	  blacklist=rs("blacklist")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub

function GuestInfo(ID,Guest,Sex)
  'Dim rs,sql
  'Set rs=server.CreateObject("adodb.recordset")
  'sql="Select * From NwebCn_Members where ID="&ID
  'rs.open sql,conn,1,1
  'if rs.bof and rs.eof then
    GuestInfo=Guest & "&nbsp;" & Sex
  'else
    'GuestInfo="<font color='green'>会员&nbsp;</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"' onClick='changeAdminFlag(""前台会员资料"")'>"&Guest&"</a>"&Sex
  'end if
  'rs.close
  'set rs=nothing
end function
%>