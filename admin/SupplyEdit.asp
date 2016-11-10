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
<TITLE>查看、修改、回复供应信息</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|96,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ReplyContent,ReplyTime,ID,NeedID,SupplyName,Remark
dim Linkman,Company,Address,ZipCode,Telephone,Fax,Mobile,Email,AddTime
ID=request.QueryString("ID")
call SupplyEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>供应信息：查看，修改，回复供应信息相关的内容</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="SupplyList.asp" onClick='changeAdminFlag("供应信息列表")'>查看供应信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>网站信息设置</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="SupplyEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">供应商品：</td>
        <td><input name="SupplyName" type="text" class="textfield" id="SupplyName" style="WIDTH: 240;" value="<%=SupplyName%>" readonly>&nbsp;<a href=NeedEdit.asp?Result=Modify&ID=<%=NeedID%> target="mainFrame" onClick='changeAdminFlag("查看需求信息")'>查看需求</a></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">补充说明：
        <td><textarea name="Remark" rows="6" class="textfield" id="Remark" style="WIDTH: 76%;" readonly><%=Remark%></textarea></td>
      </tr>
      <tr>
        <td height="20" align="right">供&nbsp;应&nbsp;者：</td>
        <td><%=Linkman%></td>
      </tr>
      <tr>
        <td height="20" align="right">单位名称：</td>
        <td><input name="Company" type="text" class="textfield" id="Company" style="WIDTH: 240;" value="<%=Company%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">通信地址：</td>
        <td><input name="Address" type="text" class="textfield" id="Address" style="WIDTH: 240;" value="<%=Address%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">邮　　编：</td>
        <td><input name="ZipCode" type="text" class="textfield" id="ZipCode" style="WIDTH: 120" value="<%=ZipCode%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">电　　话：</td>
        <td><input name="Telephone" type="text" class="textfield" id="Telephone" style="WIDTH: 240;" value="<%=Telephone%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">传　　真：</td>
        <td><input name="Fax" type="text" class="textfield" id="Fax" style="WIDTH: 120" value="<%=Fax%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">移动电话：</td>
        <td><input name="Mobile" type="text" class="textfield" id="Mobile" style="WIDTH: 120" value="<%=Mobile%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">电子邮箱：</td>
        <td><input name="Email" type="text" class="textfield" id="Email" style="WIDTH: 240" value="<%=Email%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">订购时间：</td>
        <td><input name="AddTime" type="text" class="textfield" id="AddTime" style="WIDTH: 240" value="<%=AddTime%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">回复时间：</td>
        <td><input name="ReplyTime" type="text" class="textfield" id="ReplyTime" style="WIDTH: 240" value="<%=ReplyTime%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">回复内容：</td>
        <td><textarea name="ReplyContent" rows="6" class="textfield" id="ReplyContent" style="WIDTH: 76%;"><%=ReplyContent%></textarea></td>
      </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;" ></td>
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
sub SupplyEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存回复供应信息
    set rs = server.createobject("adodb.recordset")
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_Supply where ID="&ID
      rs.open sql,conn,1,3
	  rs("ReplyContent")=StrReplace(Request.Form("ReplyContent"))
	  if not trim(request.Form("ReplyContent"))="" then
	    rs("ReplyTime")=now()
      end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑、回复供应信息！');changeAdminFlag('供应信息列表');location.replace('SupplyList.asp');</script>"
  else '提取留言信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Supply where ID="& ID
      rs.open sql,conn,1,1
	  NeedID=rs("NeedID")
	  SupplyName=rs("SupplyName")
	  Remark=ReStrReplace(rs("Remark"))
	  Linkman=GuestInfo(rs("MemID"),rs("Linkman"),rs("Sex"))
	  Company=rs("Company")
	  Address=rs("Address")
	  ZipCode=rs("ZipCode")
	  Telephone=rs("Telephone")
	  Fax=rs("Fax")
	  Mobile=rs("Mobile")
	  Email=rs("Email")
	  AddTime=rs("AddTime")
	  ReplyContent=ReStrReplace(rs("ReplyContent"))
	  ReplyTime=rs("ReplyTime")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub

function GuestInfo(ID,Guest,Sex)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    GuestInfo=Guest & "&nbsp;" & Sex
  else
    GuestInfo="<font color='green'>会员&nbsp;</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"' onClick='changeAdminFlag(""前台会员资料"")'>"&Guest&"</a>"&Sex
  end if
  rs.close
  set rs=nothing
end function 
%>