<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="CheckAdmin.asp"-->

<table name="Trans" id="Trans" width="100%" height="24" border="0" cellpadding="0" cellspacing="0" style="BORDER-BOTTOM: #333333 1px solid;font-family:宋体;font-size:12px;color: #333333;">
  <tr>
    <td width="240" nowrap >系统授权号：BC27F457E1FA34EF93</td>
    <td width="200" nowrap>管理员：<%=session("AdminName")%>[<%=session("UserName")%>]</td>
    <td width="36" nowrap>位置：</td>
    <td width="120" nowrap>后台管理首页</td>
    <td align="right" nowrap id="DateTime">
      <script> 
         setInterval("DateTime.innerHTML=new Date().toLocaleString()+' 星期'+'日一二三四五六'.charAt(new Date().getDay());",1000);
      </script>
    </td>
  </tr>
</table>