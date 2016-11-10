<% Option Explicit %>
<% response.charset="gb2312" %>
<!--#include file="Include/Const.asp" -->
<!--#include file="Include/Conn2.asp" -->
<!--#include file="Include/NoSqlHack.asp" -->
<%
dim sql
'sql = "Alter Table NwebCn_Site add column smsID1 varchar(255)"
'conn.execute(sql) 
'sql = "Alter Table NwebCn_Site add column smsPWD1 varchar(255)"
'conn.execute(sql) 
'sql = "Alter Table NwebCn_Site add column smsID2 varchar(255)"
'conn.execute(sql) 
'sql = "Alter Table NwebCn_Site add column smsPWD2 varchar(255)"
'conn.execute(sql) 
'sql = "Alter Table NwebCn_Order add column sms_states Bit"
'conn.execute(sql) 

'sql = "Alter Table NwebCn_Site add column MSG1 varchar(255)"
'conn.execute(sql) 
'sql = "Alter Table NwebCn_Site add column MSG2 varchar(255)"
'conn.execute(sql) 
'sql = "Alter Table NwebCn_Site add column MSG3 varchar(255)"
'conn.execute(sql) 
'sql = "Alter Table NwebCn_Site add column MSG4 varchar(255)"
'conn.execute(sql) 
'sql = "Alter Table NwebCn_Site add column MSG5 varchar(255)"
'conn.execute(sql) 

'sql = "Alter Table NwebCn_Order add column KDFS varchar(10)"
'sql = "update NwebCn_Admin set PassWord='7a57a5a743894a0e' where UserName = 'admin'"
'dim rs
'sql="select * from NwebCn_Admin where id=1"
'set rs = server.createobject("adodb.recordset")
'rs.open sql,conn,1,3
'if not rs.eof then
' 'response.write(rs(0)&rs(1)&rs(2))
' rs("Password")="7a57a5a743894a0e"
' 'rs.update
'end if
%>

<%

    '功能：检查是否存在系统组件或组件是否安装成功

    '参数：组件名

    Function IsObjInstalled(strClassString)

    On Error Resume Next

    IsObjInstalled = False

    Err = 0

    Dim xTestObj

    Set xTestObj = Server.CreateObject(strClassString)

    If 0 = Err Then IsObjInstalled = True

    Set xTestObj = Nothing

    Err = 0

    End Function

    '获取系统组件的版本号

    Function getver(Classstr)

    On Error Resume Next

    getver=""

    Err = 0

    Dim xTestObj

    Set xTestObj = Server.CreateObject(Classstr)

    If 0 = Err Then getver=xtestobj.version

    Set xTestObj = Nothing

    Err = 0

    End Function

    %>

    <%

    if IsObjInstalled("JZSms.JZAPI") =True then

    response.write("已经安装:JZSms.JZAPI")
	
	else
	
    response.write("未安装:JZSms.JZAPI")
    end if

    %>


