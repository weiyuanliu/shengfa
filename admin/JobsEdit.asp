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
<TITLE>编辑招聘</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|98,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,JobName,ViewFlag,JobAddress,JobNumber,Emolument,EndDate,Content,px
ID=request.QueryString("ID")
call JobsEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>招聘信息：添加，修改招聘信相关的内容</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="JobsEdit.asp?Result=Add" onClick='changeAdminFlag("添加招聘信息")'>添加招聘信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="JobsList.asp" onClick='changeAdminFlag("招聘信息列表")'>查看招聘信息</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="JobsEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">职位名称：</td>
        <td><input name="JobName" type="text" class="textfield" id="JobName" style="WIDTH: 240;" value="<%=JobName%>">&nbsp;发布：<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag or Result="Add" then response.write ("checked")%>>&nbsp;*&nbsp;不少于3个字符</td>
      </tr>
      <tr>
        <td height="20" align="right">工作地点：</td>
        <td><input name="JobAddress" type="text" class="textfield" id="JobAddress" style="WIDTH: 240;" value="<%=JobAddress%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">招聘人数：</td>
        <td><input name="JobNumber" type="text" class="textfield" id="JobNumber" style="WIDTH: 240" value="<%=JobNumber%>">&nbsp;*&nbsp;6人</td>
      </tr>
      <tr>
        <td height="20" align="right">月&nbsp;薪&nbsp;水：</td>
        <td><input name="Emolument" type="text" class="textfield" id="Emolument" style="WIDTH: 240;" value="<%=Emolument%>">&nbsp;*&nbsp;3000元/月</td>
      </tr>
	        <tr>
        <td height="20" align="right">排&nbsp;序：</td>
        <td><input name="Px" type="text" class="textfield" id="Px" style="WIDTH: 240;" value="<%=px%>">&nbsp;*&nbsp;只能填写数字</td>
      </tr>
      <tr>
        <td height="20" align="right">结束日期：</td>
        <td><input name="EndDate" type="text" class="textfield" id="EndDate" style="WIDTH: 240;" value="<% if EndDate="" then response.write (DateAdd("m",3,now())) else response.write (EndDate) end if%>" maxlength="14">&nbsp;*&nbsp;默认为3个月，可手动输入日期格式</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">信息内容：<br>
		  <img title="点击进入可视化查看、编辑环境..." src="Images/Edit.gif" width="51" height="20" style="cursor:hand" onClick="OpenDialog('../Editor/EditorDialog.html?lnk=Content&file=Editor_1.html',800,520);">
        <td><!--<textarea name="Content" rows="12" class="textfield" id="Content" style="WIDTH: 86%;" readonly><%=Content%></textarea>-->

            <textarea name="Content" rows="12" class="textfield" id="Content" style="WIDTH: 86%;" ><%=Content%></textarea><br>

换行请用"&lt;br&gt;"代替,否则用编辑器编辑</td>
      </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>   
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub JobsEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("JobName")))<3 then
      response.write ("<script language=javascript> alert('职位名称为必填项目！');history.back(-1);</script>")
      response.end
    end if
    if len(trim(request.Form("JobAddress")))="" or len(trim(request.Form("JobNumber")))="" or len(trim(request.Form("Emolument")))="" then
      response.write ("<script language=javascript> alert('""工作地点、职位数量、月薪水""名称为必填项目，且不少于2个字符！');history.back(-1);</script>")
      response.end
    end if
    if len(trim(request.Form("EndDate")))<4 then
      response.write ("<script language=javascript> alert('""结束日期""名称为必填项目！');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '创建网站管理员
	  sql="select * from NwebCn_Jobs"
      rs.open sql,conn,1,3
      rs.addnew
      rs("JobName")=trim(Request.Form("JobName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("JobAddress")=trim(Request.Form("JobAddress"))
	  rs("JobNumber")=trim(Request.Form("JobNumber"))
	  rs("Emolument")=trim(Request.Form("Emolument"))
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  rs("EndDate")=trim(Request.Form("EndDate"))
	  rs("Content")=Request.Form("Content")
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_Jobs where ID="&ID
      rs.open sql,conn,1,3
      rs("JobName")=trim(Request.Form("JobName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("JobAddress")=trim(Request.Form("JobAddress"))
	  rs("JobNumber")=trim(Request.Form("JobNumber"))
	  rs("Emolument")=trim(Request.Form("Emolument"))
	 
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  rs("EndDate")=trim(Request.Form("EndDate"))
	  rs("Content")=Request.Form("Content")
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑招聘信息！');changeAdminFlag('招聘信息列表');location.replace('JobsList.asp');</script>"
  else '提取管理员信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Jobs where ID="& ID
      rs.open sql,conn,1,1
	  JobName=rs("JobName")
	  ViewFlag=rs("ViewFlag")
	  JobAddress=rs("JobAddress")
	  JobNumber=rs("JobNumber")
	  Emolument=rs("Emolument")
	  px=Rs("Px")
	  EndDate=rs("EndDate")	  
      Content=rs("Content")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>