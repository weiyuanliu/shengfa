<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'����������������������������������������������������������������
'����������������������������������������������������������������
'�������������������տƼ���ҵ��վ����ϵͳ��LISuo����������������  ��
'����������������������������������������������������������������
' ����Ȩ���С�qisehu.com
'
'�����������������տƼ����޹�˾
'��������������Add:�Ĵ�ʡ�ɶ��ж���·������181��13¥20/21��
'����������������������������������������������������������������
'����������������������������������������������������������������
%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="˳���������޹�˾" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>�ļ�����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/html.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|33,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<body><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td>Ŀ¼���� <a href="wenjian.asp?wwwroot=../upload/picfiles/"  title="�����ϴ���ͼƬ">picfiles</a>&nbsp;|&nbsp;<a href="wenjian.asp?wwwroot=../upload/EditorFiles/" title="�༭���ϴ��ļ�">EditorFiles</a>&nbsp;|&nbsp;<a href="wenjian.asp?wwwroot=../upload/downfiles/" title="�����ļ�">downfiles</a></td>
  </tr>
</table>


<%
'ȷ�����Ŀ¼
Dim WwwRoot
    WwwRoot = Request("wwwroot")
	if wwwroot="" then wwwroot="../upload/picfiles/"

Function FsoList(Byval Path)
Dim fso,folder,fc,f,strs
dim   startNo,endNo,pageSize,TotalNo,i 
    '��ȡ��ʼ   
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
set fc = folder.files '��ü�¼����
    TotalNo=fc.Count
    '��ʾ���ͷ
    With Response
        .Write "<table width='98%' border='0' align='left' cellpadding='0' class='file' cellspacing='1' bgcolor='#cccccc'>"
        .Write "<tr><td align='center' bgcolor='#eeeeee'>ͼƬԤ��</td><td height='22' bgcolor='#eeeeee'>�ļ���</td><td align='center' bgcolor='#eeeeee'>�ļ���С</td><td align='center' bgcolor='#eeeeee'>����޸�ʱ��</td><td align='center' bgcolor='#eeeeee'>����</td></tr>"
		'#---------------
	 
'#---------------


      For each f in fc
	  if   i>=startNo   and   i<endNo   then   
	 if  mid(f.name,len(f.name)-3,len(f.name))<>".jpg" and mid(f.name,len(f.name)-3,len(f.name))<>".gif"  and mid(f.name,len(f.name)-3,len(f.name))<>".bmp" then
	strs="����ͼƬ"
	 else
	strs="<img src="&WwwRoot&F.Name&" width=""100"" height=""50"" onload='javascript:DrawImage(this,40,40);' />" 
	 end if
   '  response.Write(mid(f.name,len(f.name)-3,len(f.name)))
response.Write "  <tr><td align=center bgcolor='#eeeeee'>"&strs&"</td><td height='22' bgcolor='#eeeeee'><A HREF='"&WwwRoot&F.Name&"'>"&F.Name&"</A></td><td align='center' bgcolor='#eeeeee'>"&Round(F.Size/1024,0)&"/kb</td><td align='center' bgcolor='#eeeeee'>"&f.DateCreated&"</td> <td align='center' bgcolor='#eeeeee'>&nbsp;[<a href='wenjian.Asp?WwwRoot="&WwwRoot&"&Re=del&Ref="&WwwRoot&F.Name&"' >ɾ��</a>]</td> </tr>"
		 
		 else 
		 end if
		  i=i+1
      Next
	  response.write   "<tr><td><Hr>"   
  response.write   "<a   href=wenjian.asp>��ҳ</a>|"   
  If   startNo>1   Then   
  If   StartNo-PageSize>0   Then   
    response.write   "<a   href=wenjian.asp?startNo="&StartNo-PageSize&">��һҳ</a>|"   
      Else   
  response.write   "<a   href=wenjian.asp>��һҳ</a>|"   
  End   If   
  End   If   
  IF   endNo<TotalNo   Then   
  response.write   "<a   href=wenjian.asp?startNo="&EndNo&">��һҳ</a>|"   
  End   If   
  if   totalno mod pagesize =0 then
  response.write   "<a   href=wenjian.asp?startNo="&TotalNo-pagesize&">ĩҳ</a>"   
  else
   response.write   "<a   href=wenjian.asp?startNo="&TotalNo- (totalno mod pagesize)&">ĩҳ</a>"   
  end if

response.Write "</td></tr></table>"

    
        '  endtime=timer()   
        '  response.Write   "<br><font   color=red>ҳ��ִ��ʱ��"&FormatNumber((endtime-startime)*1000,3)&"   ����</font><br>"   
          '��������   
  If   i>endNo   then   response.end   
Set Fso=Nothing
End With
End Function

'���ú���
   Response.Write FsoList(WwwRoot)
 if Request("re")="del" THen
	 DelFile(Request("Ref"))
 end if
 Function DelFile(Files)
dim fs,file
Set fs = Server.CreateObject("Scripting.FileSystemObject")
File = Server.MapPath(Files)
on Error Resume Next
fs.DeleteFile File, True 'ǿ��ɾ��ֻ���ļ�
If Err.Number = 53 Then
Response.Write File & "�ļ������ڣ�"
Response.End
Elseif Err.Number = 70 Then
Response.Write File & "�ļ�����Ϊ����״̬��"
Response.End
Elseif Err.Number <> 0 Then
Response.Write "δ֪���󣬴�����룺" & Err.Number
Response.End
Else
Response.Write "�ɹ�ɾ���ļ���" & File
End If
response.Redirect("wenjian.Asp?WwwRoot="&WwwRoot)
End Function
   %>
   
   
</body>
</html>
