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
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|92,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ReplyContent,ReplyTime,ID,MesName,Content,ViewFlag,SecretFlag
dim Linkman,Company,Address,ZipCode,Telephone,Fax,Mobile,Email,AddTime
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
  <form name="editForm" method="post" action="MessageEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">留言主题：</td>
        <td><input name="MesName" type="text" class="textfield" id="MesName" style="WIDTH: 240;" value="<%=MesName%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">留言内容：
        <td><textarea name="Content" rows="6" class="textfield" id="Content" style="WIDTH: 76%;"><%=Content%></textarea>&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">留&nbsp;言&nbsp;者：</td>
        <td><%=Linkman%></td>
      </tr>
      <tr>
        <td height="20" align="right">单位名称：</td>
        <td><input name="Company" type="text" class="textfield" id="Company" style="WIDTH: 240;" value="<%=Company%>" ></td>
      </tr>
      <tr>
        <td height="20" align="right">通信地址：</td>
        <td><input name="Address" type="text" class="textfield" id="Address" style="WIDTH: 240;" value="<%=Address%>" ></td>
      </tr>
      <tr>
        <td height="20" align="right">邮　　编：</td>
        <td><input name="ZipCode" type="text" class="textfield" id="ZipCode" style="WIDTH: 120" value="<%=ZipCode%>" ></td>
      </tr>
      <tr>
        <td height="20" align="right">电　　话：</td>
        <td><input name="Telephone" type="text" class="textfield" id="Telephone" style="WIDTH: 240;" value="<%=Telephone%>" ></td>
      </tr>
      <tr>
        <td height="20" align="right">传　　真：</td>
        <td><input name="Fax" type="text" class="textfield" id="Fax" style="WIDTH: 120" value="<%=Fax%>" ></td>
      </tr>
      <tr>
        <td height="20" align="right">IP：</td>
        <td><input name="Mobile" type="text" class="textfield" id="Mobile" style="WIDTH: 120" value="<%=Mobile%>" ></td>
      </tr>
      <tr>
        <td height="20" align="right">电子邮箱：</td>
        <td><input name="Email" type="text" class="textfield" id="Email" style="WIDTH: 240" value="<%=Email%>" ></td>
      </tr>
      <tr>
        <td height="20" align="right">状　　态：</td>
        <td><input name="SecretFlag" type="checkbox" id="SecretFlag" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if SecretFlag then response.write ("checked")%>>
        &nbsp;首页推荐&nbsp;
        <input name="ViewFlag" type="checkbox" id="ViewFlag" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if ViewFlag then response.write ("checked")%>>&nbsp;通过审核</td>
      </tr>
      <tr>
        <td height="20" align="right">留言时间：</td>
        <td><input name="AddTime" type="text" class="textfield" id="AddTime" style="WIDTH: 240" value="<%=AddTime%>"  ></td>
      </tr>
      <tr>
        <td height="20" align="right">回复时间：</td>
        <td><input name="ReplyTime" type="text" class="textfield" id="ReplyTime" style="WIDTH: 240" value="<%=ReplyTime%>"  ></td>
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
sub MesEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("MesName")))<3 then
      response.write ("<script language=javascript> alert('""留言主题""为必填项目，且不少于3个字符！');history.back(-1);</script>")
      response.end
    end if
   
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_Message where ID="&ID
      rs.open sql,conn,1,3
      rs("MesName")=trim(Request.Form("MesName"))
      rs("Content")= StrReplace(Request.Form("Content"))
	  Rs("Company")=trim(Request.Form("Company"))
	  Rs("Address")=trim(Request.Form("Address"))
	  Rs("mobile")=trim(Request.Form("mobile"))
	  Rs("AddTime")=trim(Request.Form("AddTime"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SecretFlag")=1 then
        rs("SecretFlag")=Request.Form("SecretFlag")
	  else
        rs("SecretFlag")=0
	  end if
	  rs("ReplyContent")=StrReplace(Request.Form("ReplyContent"))
	  if not (trim(request.Form("ReplyContent"))="" or trim(request.Form("ReplyTime"))<>"") then
	    rs("ReplyTime")=now()
      end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功审核、编辑、回复留言信息！');changeAdminFlag('留言信息列表');location.replace('MessageList.asp');</script>"
  else '提取留言信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Message where ID="& ID
      rs.open sql,conn,1,1
	  MesName=rs("MesName")
	  Content=ReStrReplace(rs("Content"))
	  Linkman=GuestInfo(rs("MemID"),rs("Linkman"),rs("Sex"))
	  Company=rs("Company")
	  Address=rs("Address")
	  ZipCode=rs("ZipCode")
	  Telephone=rs("Telephone")
	  Fax=rs("Fax")
	  Mobile=rs("Mobile")
	  Email=rs("Email")
	  ViewFlag=rs("ViewFlag")
	  SecretFlag=rs("SecretFlag")
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