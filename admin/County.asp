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
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>二级市级信息管理</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AddCounty.asp?Result=Add" onClick='changeAdminFlag("添加三级区信息")'>添加三级区信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="County.asp" onClick='changeAdminFlag("查看所有三级区信息")'>查看所有三级区信息</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=County" method="post" name="formDel" >
    <tr>
      <td width="79" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID编号</strong></font></td>
      <td width="92" align="center"  bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>所属省份</strong></font></td>
      <td width="81" height="24" align="center"  bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>市名称</strong></font></td>
      <td width="112" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>区名称</strong></font></td>
      <td width="84" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>排列顺序</strong></font></td>
      <td width="113" align="center"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">添加时间</font></strong></td>
      <td width="90" align="center" bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">      </td>
    </tr>
	<%Call CityList(20) %>
  </form>
</table>
</body>
</html>
<%
Sub CityList(Page_Size)
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from County order by px asc,id asc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='7'>")
				response.Write("暂无信息！")
			response.Write("</td>")
		response.Write("</tr>")
	else
		rs.pagesize=page_size
		dim sum_page,total,i
		total=rs.recordcount
		sum_page=total \ page_size
		if total mod page_size <>0 then sum_page=sum_page+1
		dim page
		page=trim(request.querystring("page"))
		if page="" or isnull(page) or (not IsNumeric(page)) then
			page=1
		elseif Cint(Page)<=1 then
			page=1
		elseif Cint(page) => sum_page then
			page=sum_page
		else
			page=Cint(page)
		end if
		rs.absolutepage=page
		
		for i=1 to Page_Size 
			if not rs.eof then
				response.Write("<tr bgcolor='#EBF2F9'>")
					
					response.Write("<td align='center'>")
						response.Write(rs("ID"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(GetValue(rs("ParentID"),"Province","Content"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						Response.Write(GetValue(rs("ParentID2"),"City","Content"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("Content"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("Px"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("AddTime"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write("【<a href='EditCounty.asp?ID="&rs("ID")&"'>编 辑</a>】")
						response.Write("<input type='checkbox' name='SelectID' id='SelectID' value='"&rs("ID")&"' style='margin-left:5px;'>选 择")
					response.Write("</td>")
				response.Write("</tr>")
				rs.movenext
			end if
		next
		
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='6'></td>")
			response.Write("<td align='center'>")
				response.Write("<input name='DelRecord' type='submit' value='删 除'>")
			response.Write("</td>")
		response.Write("</tr>")
		
		if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
	end if
	rs.close()
	set rs=Nothing
End sub
%>
<%
sub Contrl_Page(page,sum_page,total,page_size) 
dim Url,linkfile,pagewhere,UrlValue
Url=request.ServerVariables("URL")
Url=mid(Url,InstrRev(Url,"/")+1)
linkfile=Url
UrlValue=""

if UrlValue<>"" then
	pagewhere=UrlValue
end if

	response.Write("<tr bgcolor='#EBF2F9'>")
		response.Write("<td colspan='5' class='Item_list' style='padding-top:5px; padding-bottom:5px; text-align:right;'>")
			response.Write("[共计："&total&"条] ")
					response.write("[每页："&page_size&"条] ")
					response.write("[页次："&page&"/"&sum_page&"] ")
					if page<=1 then
						response.write("[首页] [上一页] ")
					else 
						response.write("<a href='"&linkfile&"?page=1"&pagewhere&"'>")
						response.write("[首页]")
						response.write("</a> ")
						response.write("<a href='"&linkfile&"?page="&page-1&pagewhere&"'>")
						response.write("[上一页]")
						response.write("</a> ")
					end if
					
					if page < sum_page then
						response.write("<a href='"&linkfile&"?page="&page+1&pagewhere&"'>")
						response.write("[下一页]")
						response.write("</a> ")
					else
						response.write("[下一页] ")
					end if
					
					if sum_page>1 and page < sum_page then
						response.write("<a href='"&linkfile&"?page="&sum_page&pagewhere&"'>")
						response.write("[末页]")
						response.write("</a>")
					else
						response.write("[末页]")
					end if
					dim cc
					response.write(" 转到：")%>
					<select name="page" size="1" onChange="javascript:window.location='<%=linkfile%>?page='+this.options[this.selectedIndex].value+'<%=pagewhere%>';">
						<%for cc=1 to sum_page
							if cc=page then
								response.write("<option value='"&cc&"' selected >"&cc&"页")
							else
								response.write("<option value='"&cc&"'>"&cc&"页")
							end if
						next%>
					</select>
		<%response.Write("</td>")
	response.Write("</tr>")
end sub

function  GetValue(ParentID,tablename,Content)
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select "&Content&" from "&tablename&" where id="&ParentID
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		GetValue=rs(Content)		
	end if
	rs.close()
	set rs=Nothing
end function
%>

