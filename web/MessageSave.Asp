<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　　七日科技企业网站管理系统（LiSuo）　　　　　　　  ┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'   版权所有　qisehu.com
'   程序制作　七日科技网络有限公司
'　 Add:四川省成都市二环路西三段181号13楼20/21号
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%>
<% Option Explicit %>
<% response.charset="gb2312" %>
<!--#include file="../Include/NoSqlHack.asp" -->
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
dim rs,sql,SiteTitle,SiteUrl,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,IcpNumber,MesViewFlag,syimg,gonggao,ybpz,qq,syjs
set rs = server.createobject("adodb.recordset")
sql="select top 1 * from NwebCn_Site"
rs.open sql,conn,1,1
SiteTitle=rs("SiteTitle")
SiteUrl=rs("SiteUrl")
ComName=rs("ComName")
Address=rs("Address")
ZipCode=rs("ZipCode")
Telephone=rs("Telephone")
Fax=rs("Fax")
Email=rs("Email")
Keywords=rs("Keywords")
Descriptions=rs("Descriptions")
IcpNumber=rs("IcpNumber")
MesViewFlag=rs("MesViewFlag")
syimg=rs("syimg")
gonggao=rs("Gonggao")
ybpz=rs("ybpz")
QQ=RS("QQ")
syjs=rs("syjs")
rs.close
set rs=nothing '
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=siteTitle%></title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<META NAME="Keywords" CONTENT="<% =Keywords %>" />
<META NAME="Description" CONTENT="<% =Descriptions %>" />
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #010101;
}
body,td,th {
	font-size: 12px;
	color: #ababab;
}
.STYLE1 {
	font-size: 14px;
	font-weight: bold;
	color: #CDD206;
}
-->
</style></head>

<body>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="Images/Top.jpg" width="1000" height="111" /></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/dhbg.jpg">
  <tr>
    <td><img src="Images/dhleft.jpg" width="63" height="55" /></td>
    <td width="150" align="center"><a href="Index.Asp"><img src="Images/wzsy.jpg" width="81" height="55" border="0" /></a></td>
    <td width="17" align="center"><img src="Images/dhfg.jpg" width="17" height="55" /></td>
    <td width="150" align="center"><a href="Cpsm.Asp" target="_blank"><img src="Images/dhcpsm.jpg" width="91" height="55" border="0" /></a></td>
    <td width="17" align="center"><img src="Images/dhfg.jpg" width="17" height="55" /></td>
    <td width="144" align="center"><a href="LiuYan.Asp" target="_blank"><img src="Images/dhkfly.jpg" width="85" height="55" border="0" /></a></td>
    <td width="17" align="center"><img src="Images/dhfg.jpg" width="17" height="55" /></td>
    <td width="144" align="center"><a href="Wtjd.Asp" target="_blank"><img src="Images/dhwtjd.jpg" width="88" height="55" border="0" /></a></td>
    <td width="17" align="center"><img src="Images/dhfg.jpg" width="17" height="55" /></td>
    <td width="150" align="center"><a href="../zxdg.asp" ><img src="Images/dhzxdg.jpg" width="90" height="55" border="0" /></a></td>
    <td width="62" align="right"><img src="Images/dhright.jpg" width="45" height="55" /></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#000000" style="padding-top:2px; padding-bottom:2px"><img src="Images/banner.jpg" width="1000" height="213" /></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#010101">
  <tr>
    <td width="300" valign="top"><!--#include file="left.asp"--></td>
    <td width="700" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td align="center"><img src="Images/nrongzj.jpg" width="687" height="58" /></td>
      </tr>
      <tr>
        <td><img src="Images/t1.jpg" width="687" height="147" /></td>
      </tr>
      <tr>
        <td><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td style="line-height:22px"><%=MessageSave()%><%
Function MessageSave()
if Request.ServerVariables("HTTP_REFERER")="" then
	response.write("对不起！数据出错，请<a href='#' onclick='window.location.href='index.asp''>返回！</a>")
	response.end
end if
dim MesName,Content,SecretFlag,MemID,Linkman,Sex,Company,Address,ZipCode,Telephone,Fax,Mobile,Email,VerifyCode
dim rs,sql
MesName=trim(request.form("MesName"))
Content=request.form("Content")
if trim(request.form("SecretFlag"))="1" then
  SecretFlag=1
else
  SecretFlag=0
end if
MemID=0
Linkman=trim(request.form("Linkman"))
 

 
 
 
set rs = server.createobject("adodb.recordset")
sql="select * from NwebCn_Message"
rs.open sql,conn,1,3
rs.addnew
rs("MesName")=StrReplace(MesName)
rs("Content")=StrReplace(Content)
rs("MemID")=MemID
rs("Linkman")=StrReplace(Linkman)
rs("Sex")=Sex
Rs("Mobile")=Userip() 

rs("SecretFlag")=SecretFlag
rs("AddTime")=now()
rs.update
rs.close
set rs=nothing
Response.Write("<script language='javascript'>alert('感谢您的留言，留言内容要通过我们工作人员审核后才能发表。'); window.location.href('liuyan.asp');</script>")

End Function
'****************************** 
Function Userip() 
Dim GetClientIP 
'如果客户端用了代理服务器，则应该用ServerVariables("HTTP_X_FORWARDED_FOR")方法 
GetClientIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If GetClientIP = "" or isnull(GetClientIP) or isempty(GetClientIP) Then 
'如果客户端没用代理，应该用Request.ServerVariables("REMOTE_ADDR")方法 
GetClientIP = Request.ServerVariables("REMOTE_ADDR") 
end if 
Userip = GetClientIP 
End function 
%></td>
                                  </tr>
          </table></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      
    </table></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/downbg.jpg">
  <tr>
    <td width="34" rowspan="2"><img src="Images/downleft.jpg" width="34" height="98" /></td>
    <td width="933" height="26" align="center">版权所有 &copy; 倍洛加中国官方网站 地址:<%=Address%>ICP备案：粤ICP备06121108号</td>
    <td width="33" rowspan="2"><img src="Images/downright.jpg" width="33" height="98" /></td>
  </tr>
  <tr>
    <td align="center"><table width="80%%" border="0" cellpadding="4" cellspacing="0">
        <tr>
          <td><img src="Images/bottom_aaa.gif" width="120" height="36"></td>
          <td><img src="Images/bottom_other3.gif" width="120" height="36"></td>
          <td><img src="Images/bottom_wenmingchengxin.gif" width="120" height="36"></td>
          <td><img src="Images/cool.gif" width="120" height="36"></td>
          <td><table width="90%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td width="42%"><img src="Images/ind36.gif" width="35" height="43"></td>
                <td width="58%" style="line-height:18px">经营性网站<br>
                备案信息</td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<div style="display:none"><script src="http://s11.cnzz.com/stat.php?id=1109194&web_id=1109194&show=pic" language="JavaScript" charset="gb2312"></script>
</div>
</body>
</html>
