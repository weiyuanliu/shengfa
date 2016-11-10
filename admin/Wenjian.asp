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
<META NAME="Author" CONTENT="顺意网络有限公司" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>文件操作</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/html.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|33,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<body><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td>目录管理 <a href="wenjian.asp?wwwroot=../upload/picfiles/"  title="单独上传得图片">picfiles</a>&nbsp;|&nbsp;<a href="wenjian.asp?wwwroot=../upload/EditorFiles/" title="编辑器上传文件">EditorFiles</a>&nbsp;|&nbsp;<a href="wenjian.asp?wwwroot=../upload/downfiles/" title="下载文件">downfiles</a></td>
  </tr>
</table>


<%
'确定浏览目录
Dim WwwRoot
    WwwRoot = Request("wwwroot")
	if wwwroot="" then wwwroot="../upload/picfiles/"

Function FsoList(Byval Path)
Dim fso,folder,fc,f,strs
dim   startNo,endNo,pageSize,TotalNo,i 
    '提取开始   
    if   Int(Request.QueryString("startNo"))<1   then   
  startNo=0  
    else   
  startNo=Int(Request.QueryString("startNo"))   
    end   if   
    pageSize=cInt(40)   
	
    endNo=Int(startNo)+cInt(pageSize) 
	 
    i=0   
	 
set Fso=createobject("Scripting.filesystemobject")

set folder = fso.getfolder(Server.MapPath(Path))
set fc = folder.files '获得记录总数
    TotalNo=fc.Count
    '显示表格头
    With Response
        .Write "<table width='98%' border='0' align='left' cellpadding='0' class='file' cellspacing='1' bgcolor='#cccccc'>"
        .Write "<tr><td align='center' bgcolor='#eeeeee'>图片预览</td><td height='22' bgcolor='#eeeeee'>文件名</td><td align='center' bgcolor='#eeeeee'>文件大小</td><td align='center' bgcolor='#eeeeee'>最后修改时间</td><td align='center' bgcolor='#eeeeee'>操作</td></tr>"
		'#---------------
	 
'#---------------


      For each f in fc
	  if   i>=startNo   and   i<endNo   then   
	 if  mid(f.name,len(f.name)-3,len(f.name))<>".jpg" and mid(f.name,len(f.name)-3,len(f.name))<>".gif"  and mid(f.name,len(f.name)-3,len(f.name))<>".bmp" then
	strs="不是图片"
	 else
	strs="<img src="&WwwRoot&F.Name&" width=""100"" height=""50"" onload='javascript:DrawImage(this,40,40);' />" 
	 end if
   '  response.Write(mid(f.name,len(f.name)-3,len(f.name)))
response.Write "  <tr><td align=center bgcolor='#eeeeee'>"&strs&"</td><td height='22' bgcolor='#eeeeee'><A HREF='"&WwwRoot&F.Name&"'>"&F.Name&"</A></td><td align='center' bgcolor='#eeeeee'>"&Round(F.Size/1024,0)&"/kb</td><td align='center' bgcolor='#eeeeee'>"&f.DateCreated&"</td> <td align='center' bgcolor='#eeeeee'>&nbsp;[<a href='wenjian.Asp?WwwRoot="&WwwRoot&"&Re=del&Ref="&WwwRoot&F.Name&"' >删除</a>]</td> </tr>"
		 
		 else 
		 end if
		  i=i+1
      Next
	  response.write   "<tr><td><Hr>"   
  response.write   "<a   href=wenjian.asp>首页</a>|"   
  If   startNo>1   Then   
  If   StartNo-PageSize>0   Then   
    response.write   "<a   href=wenjian.asp?startNo="&StartNo-PageSize&">上一页</a>|"   
      Else   
  response.write   "<a   href=wenjian.asp>上一页</a>|"   
  End   If   
  End   If   
  IF   endNo<TotalNo   Then   
  response.write   "<a   href=wenjian.asp?startNo="&EndNo&">下一页</a>|"   
  End   If   
  if   totalno mod pagesize =0 then
  response.write   "<a   href=wenjian.asp?startNo="&TotalNo-pagesize&">末页</a>"   
  else
   response.write   "<a   href=wenjian.asp?startNo="&TotalNo- (totalno mod pagesize)&">末页</a>"   
  end if

response.Write "</td></tr></table>"

    
        '  endtime=timer()   
        '  response.Write   "<br><font   color=red>页面执行时间"&FormatNumber((endtime-startime)*1000,3)&"   毫秒</font><br>"   
          '如果则结束   
  If   i>endNo   then   response.end   
Set Fso=Nothing
End With
End Function

'调用函数
   Response.Write FsoList(WwwRoot)
 if Request("re")="del" THen
	 DelFile(Request("Ref"))
 end if
 Function DelFile(Files)
dim fs,file
Set fs = Server.CreateObject("Scripting.FileSystemObject")
File = Server.MapPath(Files)
on Error Resume Next
fs.DeleteFile File, True '强制删除只读文件
If Err.Number = 53 Then
Response.Write File & "文件不存在！"
Response.End
Elseif Err.Number = 70 Then
Response.Write File & "文件属性为锁定状态！"
Response.End
Elseif Err.Number <> 0 Then
Response.Write "未知错误，错误编码：" & Err.Number
Response.End
Else
Response.Write "成功删除文件！" & File
End If
response.Redirect("wenjian.Asp?WwwRoot="&WwwRoot)
End Function
   %>
   
   
</body>
</html>
