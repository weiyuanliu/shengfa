<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Images/CssAdmin.css">
<title>�ļ�ѡ��</title>
</head> 
<!--#include file="CheckAdmin.asp"-->
<body>
<table width="400" border="0" align="center" cellpadding="12" cellspacing="1" bgcolor="#6298E1">
  <form action="UpFileSave.asp?Result=<%=request.QueryString("Result")%>" method="post" enctype="multipart/form-data" name="formUpload">
  <tr>
    <td bgcolor="#EBF2F9">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="60" height="30" nowrap>ѡ���ļ���</td>
        <td><input name="FromFile" type="file" class="textfield" id="FromFile" size="41"></td>
      </tr>
      <tr>
        <td height="30">�ϴ�λ�ã�</td>
        <td><select name="SaveToPath" class="textfield">
          <option value="../Upload/PicFiles/" selected>ͼƬ�ļ� /Upload/PicFiles</option>
          <option value="../Upload/DownFiles/">�����ļ� /Upload/DownFiles</option>
          <option value="../Upload/OtherFiles/">�����ļ� /Upload/OtherFiles</option>
        </select></td>
      </tr>
      <tr>
        <td height="36" colspan="2" align="center" valign="bottom"><input name="reset" type="reset" class="button" value=" ���� ">
          &nbsp;<input name="Submit" type="submit" class="button" value=" �ϴ� "></td>
        </tr>
    </table>
	</td>
  </tr>
  </form>
</table>
</body>
</html>

