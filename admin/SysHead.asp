<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="CheckAdmin.asp"-->

<table name="Trans" id="Trans" width="100%" height="24" border="0" cellpadding="0" cellspacing="0" style="BORDER-BOTTOM: #333333 1px solid;font-family:����;font-size:12px;color: #333333;">
  <tr>
    <td width="240" nowrap >ϵͳ��Ȩ�ţ�BC27F457E1FA34EF93</td>
    <td width="200" nowrap>����Ա��<%=session("AdminName")%>[<%=session("UserName")%>]</td>
    <td width="36" nowrap>λ�ã�</td>
    <td width="120" nowrap>��̨������ҳ</td>
    <td align="right" nowrap id="DateTime">
      <script> 
         setInterval("DateTime.innerHTML=new Date().toLocaleString()+' ����'+'��һ����������'.charAt(new Date().getDay());",1000);
      </script>
    </td>
  </tr>
</table>