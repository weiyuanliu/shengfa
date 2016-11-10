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
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>编辑新闻</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<%
call CreateEditor("Content")
%>

</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|23,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,NewsName,ViewFlag,SortName,SortID,SortPath
dim GroupID,GroupIdName,Exclusive,NoticeFlag,Source,Content,px,CommendFlag,smallpic,bigpic,daodu
ID=request.QueryString("ID")
call NewsEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>新闻检索及分类查看：添加，修改，删除新闻信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="NewsEdit.asp?Result=Add" onClick='changeAdminFlag("添加产品信息")'>添加新闻信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NewsList.asp" onClick='changeAdminFlag("产品列表")'>查看所有新闻信息</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="NewsEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">新闻名称：</td>
        <td><input name="NewsName" type="text" class="textfield" id="NewsName" style="WIDTH: 240;" value="<%=NewsName%>" maxlength="100">&nbsp;显示：<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag or Result="Add" then response.write ("checked")%>>&nbsp;*&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">所属类别：</td>
        <td><input name="SortName" type="text" class="textfield" id="SortName" value="<%=SortName%>" style="WIDTH: 240;background-color:#EBF2F9;" readonly>&nbsp;<a href="javaScript:OpenScript('SelectSort.asp?Result=News',500,500,'')"><img src="Images/Select.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">类别数字：</td>
        <td><input name="SortID" type="text" class="textfield" id="SortID" style="WIDTH: 40;background-color:#EBF2F9;" value="<%=SortID%>" readonly><input name="SortPath" type="text" class="textfield" id="SortPath" style="WIDTH: 200;background-color:#EBF2F9;" value="<%=SortPath%>" readonly>&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">新闻来源：</td>
        <td><input name="Source" type="text" class="textfield" style="WIDTH: 240;" value="<%=Source%>" maxlength="100"></td>
      </tr>
	    <tr>
        <td height="20" align="right">排序：</td>
        <td><input name="px" type="text" class="textfield" style="WIDTH: 60px;" value="<%=px%>" maxlength="100">*只能输入数字</td>
      </tr>

      <tr>
        <td height="20" align="right">标　　记：</td>
        <td><input name="NoticeFlag" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if NoticeFlag then response.write ("checked")%>>&nbsp;公告&nbsp;&nbsp;<input name="CommendFlag" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if CommendFlag then response.write ("checked")%>>&nbsp;推荐</td>
      </tr>
	        <tr>
        <td height="20" align="right">新闻主图：</td>
        <td><input name="BigPic" type="text" class="textfield" style="WIDTH: 240;" value="<%=BigPic%>" maxlength="100">
        &nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=BigPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"> 推荐341*199 </a><a href="javaScript:OpenScript('UpFileForm.asp?Result=SmallPic',460,180)">幻灯图片必须是jpg格式</a></td>
      </tr>
      <tr>
        <td height="20" align="right">新闻缩 略 图：</td>
        <td><input name="SmallPic" type="text" class="textfield" style="WIDTH: 240;" value="<%=SmallPic%>" maxlength="100">
        &nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=SmallPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"> 推荐130*85 </a></td>
      </tr>
      <tr>
        <td height="20" align="right">查看权限：</td>
        <td><select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> 隶属<input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
      </tr>
	  	    <tr>
        <td height="20" align="right">导　　读：</td>
        <td> <input  name="daodu" size='120' class="textfield"  value="<%=daodu%>" />        不超过200字符</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">信息内容：<br>
        <td style="padding:6px"><textarea name="Content" rows="30" class="textfield" id="Content" style="WIDTH: 86%;" ><%=Content%></textarea></td>
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
sub NewsEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑产品信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("NewsName")))<1 then
      response.write ("<script language=javascript> alert('新闻名称为必填项目！');history.back();</script>")
      response.end
    end if
    if Result="Add" then '创建产品信息
	  sql="select * from NwebCn_News"
      rs.open sql,conn,1,3
      rs.addnew
      rs("NewsName")=trim(Request.Form("NewsName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")="" and Request.Form("SortPath")="" then
        response.write ("<script language=javascript> alert('请选择所属分类！');history.back();</script>")
        response.end
	  else
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  end if
	  rs("Source")=trim(Request.Form("Source"))
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  if Request.Form("NoticeFlag")=1 then
        rs("NoticeFlag")=Request.Form("NoticeFlag")
	  else
        rs("NoticeFlag")=0
	  end if
	  if Request.Form("CommendFlag")=1 then
        rs("CommendFlag")=Request.Form("CommendFlag")
	  else
        rs("CommendFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("Content")=trim(Request.Form("Content"))
	  rs("BigPic")=trim(Request.Form("BigPic"))	  
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  rs("daodu")=trim(Request.Form("daodu"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改产品信息
      sql="select * from NwebCn_News where ID="&ID
      rs.open sql,conn,1,3
      rs("NewsName")=trim(Request.Form("NewsName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")<>"" and Request.Form("SortPath")<>"" then
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  else
        response.write ("<script language=javascript> alert('请选择所属分类！');history.back();</script>")
        response.end
	  end if
	  rs("Source")=trim(Request.Form("Source"))
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  if Request.Form("NoticeFlag")=1 then
        rs("NoticeFlag")=Request.Form("NoticeFlag")
	  else
        rs("NoticeFlag")=0
	  end if
	   if Request.Form("commendFlag")=1 then
        rs("commendFlag")=Request.Form("commendFlag")
	  else
        rs("commendFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("Content")=trim(Request.Form("Content")) 
	  rs("BigPic")=trim(Request.Form("BigPic"))	  
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  
	  rs("daodu")=trim(Request.Form("daodu"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑新闻信息！');changeAdminFlag('新闻列表');location.replace('NewsList.asp');</script>"
  else '提取产品信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_News where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("数据库读取记录出错！")
        response.end
      end if
	  NewsName=rs("NewsName")
	  ViewFlag=rs("ViewFlag")
	  SortName=SortText(rs("SortID"))
	  SortID=rs("SortID")
	  SortPath=rs("SortPath")
	  Source=rs("Source")
	  px=rs("px")
	  NoticeFlag=rs("NoticeFlag")
	  CommendFlag=rs("CommendFlag")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
      Content=rs("Content")
	  BigPic=rs("BigPic")
	   daodu=rs("daodu")
	  SmallPic=rs("SmallPic")
	  rs.close
      set rs=nothing 
    end if
  end if
end sub
%>
<% 
sub SelectGroup()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from NwebCn_MemGroup"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("未设组别")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")
    if GroupID=rs("GroupID") then response.write ("selected")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext
  wend
  rs.close
  set rs=nothing
end sub
%>
<%
'生成所属类别--------------------------
Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_NewsSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>