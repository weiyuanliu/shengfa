<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Server.scriptTimeout=300 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����ļ�</title>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="JavaScript"> 
<!-- 
function CopyPath(FilePath,FileSize)
{
    var str=location.toString()
    var Result=((((str.split('?'))[1]).split('='))[1]);
	window.opener.editForm(Result).focus();								
	window.opener.document.editForm(Result).value=FilePath;
    if (Result == "FileUrl")
	{
	   window.opener.document.editForm.FileSize.value=FileSize;
    }
	window.opener=null;
    window.close();
}
//--> 
</script> 
</head>
<!--#include file="UpFileClass.asp"-->
<body >
<table width=400 border=0 align="center" cellpadding="12" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td width=100% height=100% align="center" bgcolor="#EBF2F9" class=tablebody1 >
<%
dim Upload,File,FormName,SaveToPath,FileName,FileExt
dim RanNum
call UpFile()
'===========������ϴ�(upload_0)====================
sub UpFile()
  set Upload=new UpFile_Class '�����ϴ�����
  Upload.GetData (1024*1024*15) 'ȡ���ϴ�����,�˴���Ϊ15M

  if Upload.err > 0 then
    select case Upload.err
      case 1
        Response.Write "����ѡ����Ҫ�ϴ����ļ���<a href=# onclick=history.go(-1)>����</a>&nbsp;��"
      case 2
        Response.Write "�ļ���С����������15M��<a href=# onclick=history.go(-1)>����</a>&nbsp;��"
    end select
    exit sub
  else
    SaveToPath=Upload.form("SaveToPath") '�ļ�����Ŀ¼,��Ŀ¼����Ϊ����ɶ�д
    if SaveToPath="" then
      SaveToPath="../"
    end if
    '��Ŀ¼���(/)
    if right(SaveToPath,1)<>"/" then 
      SaveToPath=SaveToPath&"/"
    end if 
    for each FormName in Upload.file '�г������ϴ��˵��ļ�
      set file=Upload.file(FormName) '����һ���ļ�����
      if file.Filesize<100 then
        response.write "����ѡ����Ҫ�ϴ����ļ���<a href=# onclick=history.go(-1)>����</a>&nbsp;��"
        response.end
      end if

      FileExt=lcase(File.FileExt)
      if CheckFileExt(FileEXT)=false then
        response.write "�ļ���ʽ�������ϴ���<a href=# onclick=history.go(-1)>����</a>&nbsp;��"
        response.end
      end if

      randomize timer
      RanNum=int(9000*rnd)+1000
      Filename=SaveToPath&year(now)&"."&month(now)&"."&day(now)&"_"&hour(now)&"."&minute(now)&"."&Second(now)&"_"&RanNum&"."&fileExt
      if file.FileSize>0 then '��� FileSize > 0 ˵�����ļ�����
        Result=file.SaveToFile(Server.mappath(FileName)) '�����ļ�
        if Result="ok" then
		
		  response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
          response.write "<tr>"
          response.write "<td width='60' height='30'>�ϴ��ɹ���</td>"
          response.write "<td nowrap><font color='#ff0000'>"&File.FilePath&file.FileName&"</font></td>"
          response.write "</tr>"
          response.write "<tr>"
          response.write "<td nowrap height='30'>����·����</td>"
          response.write "<td nowrap><input type='text' size='56' class='textfield' value='"&right(FileName,len(FileName))&"'></td>"
          response.write "</tr>"
          response.write "<tr>"
          response.write "<td nowrap height='30'>�ļ���С��</td>"
          response.write "<td nowrap><input type='text' size='56' class='textfield' value='"&GainFileSize(file.Filesize)&"'></td>"
          response.write "</tr>"		  
          response.write "<tr>"
         ' response.write "<td height='36' colspan='2' valign='bottom' align='center'><input name='CopyPath' type='button' class='button' value='�����ļ�·��'  onclick=""CopyPath('"&right(FileName,len(FileName))&"','"&GainFileSize(file.Filesize)&"')""></td>"
          Response.Write("<script language=javascript>CopyPath('"&right(FileName,len(FileName))&"','"&GainFileSize(file.Filesize)&"');</script>")
		  
		  response.write "</tr>"
          response.write "</table>"
        else
          response.write File.FilePath&file.FileName&"�ϴ�ʧ��&nbsp;��"&Result&"<br>"
        end if
      end if
      set file=nothing
    next
    set Upload=nothing
  end if
end sub

'�ж��ļ������Ƿ�ϸ�
Private Function CheckFileExt (FileEXT)
  dim ForumUpload
  ForumUpload="exe,gif,jpg,jpeg,rar,zip,doc,pdf"
  ForumUpload=split(ForumUpload,",")
  for i=0 to ubound(ForumUpload)
    if lcase(FileEXT)=lcase(trim(ForumUpload(i))) then
      CheckFileExt=true
      exit Function
    else
      CheckFileExt=false
    end if
  next
End Function

Private Function GainFileSize (SizeByte)
  if SizeByte < 1024*1024 then
    GainFileSize=round(SizeByte/1024,2) & "&nbsp;KB"
  else  
    GainFileSize=round(SizeByte/1024/1024,2) & "&nbsp;MB"
  end if
End Function

%>
    </td>
  </tr>
</table>
</body>
</html>