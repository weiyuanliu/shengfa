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
<TITLE>管理员列表</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
Dim Action
Action=Trim(Request.QueryString("Action"))
if Action="AddProvince" then Call SaveProvince()
%>

<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>添加二级市级信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AddCity.asp?Result=Add" onClick='changeAdminFlag("添加市级信息")'>添加市级信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="City.asp" onClick='changeAdminFlag("查看所有市级信息")'>查看所有市级信息</a></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9">
    <table width="78%" border="0" cellpadding="5" cellspacing="0">
       <form name="AddProvince" id="AddProvince" method="post" action="AddCity.asp?Action=AddProvince" onSubmit="return CheckValue();">
      <tr>
        <td height="25" align="right">所属省级：</td>
        <td height="25">
        	<select name="ParentID" id="ParentID" size="1">
            	<%Call GetParent()%>
            </select>        	　<span class="STYLE1">*必填</span> </td>
      </tr>
      <tr>
        <td width="21%" height="25" align="right">市级名称：</td>
        <td width="79%" height="25"><input type="text" name="Content" id="Content">　
          <span class="STYLE1">*必填</span></td>
      </tr>
      <tr>
        <td height="25" align="right">添加时间：</td>
        <td height="25"><input name="AddTime" type="text" id="AddTime" value="<%=now()%>"></td>
      </tr>
      <tr>
        <td height="25" align="right">排列顺序：</td>
        <td height="25"><input name="Px" type="text" id="Px" value="0"> 　
          *请填写数字，越小排在越前</td>
      </tr>
      
      <tr>
        <td height="40" align="right">&nbsp;</td>
        <td height="40"><label>
          <input type="submit" name="button" id="button" value="添 加" style="margin-right:10px;">
          <input type="reset" name="button2" id="button2" value="重 置">
        </label></td>
      </tr>
        </form>
    </table>
  
    </td>    
  </tr>
</table>
</body>
</html>
<%
sub SaveProvince()
	dim Content,AddTime,Px,ID,ParentID
	ID=Trim(Request.Form("ID"))
	Content=Trim(Request.Form("Content"))
	AddTime=Trim(Request.Form("AddTime"))
	Px=Trim(Request.Form("Px"))
	ParentID=Trim(Request.Form("ParentID"))
	
	if Content="" or isnull(Content) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('请填写省份信息！');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
	end if
	
	if Px="" or not(IsNumeric(Px)) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('请填写有效的排序信息！');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
	end if
	
	if ParentID="" or isnull(ParentID) or ParentID="Null" then
		response.Write("<script language=javascript>"&vbcrlf) 
			response.Write("alert('没有省级信息，不能添加二级信息！请先添加省级信息！');"&vbcrlf)
			response.Write("window.location.href='AddProvince.asp?Result=Add';")
	 	response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	else
		if Not(IsNumeric(ParentID)) or isnull(ParentID) or ParentID="" then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('数据出错，不能继续！');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		end if
	end if
	
	Dim rs,sql
	set rs=server.CreateObject("Adodb.Recordset")
	
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		sql="select * from City where Content='"&Content&"'"
	else
		sql="select * from City where Content='"&Content&"' and id not in("&id&")"
	end if
	rs.open sql,conn,1,1
	
	if not rs.eof and not rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script langauge=javascript>"&vbcrlf)
			response.Write("alert('对不起，不能重复添加！');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	end if
	
	rs.close()
	
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		Sql="Select top 1 * from City"
	else
		Sql="Select top 1 * from City where id="&ID
	end if
	rs.open sql,conn,1,3
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		rs.addnew()
		rs("Content")=Content
		if AddTime<>"" then
			rs("AddTime")=AddTime
		else
			rs("AddTime")=Now()
		end if
		rs("ParentID")=ParentID
		rs("Px")=Px
		rs.update()
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('操作成功！');"&vbcrlf)
			response.Write("window.location.href=document.referrer;")
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	else
		if rs.eof and rs.bof then
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('记录未找到，操作失败！');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		else
			rs("Content")=Content
			if AddTime<>"" then
				rs("AddTime")=AddTime
			else
				rs("AddTime")=Now()
			end if
			rs("ParentID")=ParentID
			rs("Px")=Px
			rs.update()
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert(更新记录成功！);"&vbcrlf)
				response.Write("window.location.href=document.referrer;")
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		end if
	end if
end sub

sub GetParent()
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from Province order by px asc,id asc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<option value='Null'>暂无信息</option>")	
	else
		while not rs.eof 
			response.Write("<option value='"&rs("id")&"'>"&rs("Content")&"</option>")
			rs.movenext
		wend
	end if
	rs.close()
	set rs=Nothing
End sub
%>

<script language="javascript">
<!--
	function CheckValue()
	{
		var Content,Px;
		Content=document.getElementById("Content");
		Px=document.getElementById("Px");
		
		if(Content.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写省份名称！");
			Content.focus();
			return false;
		}
		
		if((Px.value).search("^-?\\d+(\\.\\d+)?$")!=0)
		{
			alert("请填写有效的数字！");
			Px.select();
			return false;
		}
		return true;
	}

-->
</script>
