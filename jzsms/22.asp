<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
</head>

<body>
<% 
dim name1,phone
phone=request("phone")
name1=request("name")
Response.Write "号码"&phone
Response.Write "<br>"
Response.Write "姓名"&name1

Dim obj,dxname
set obj = server.createobject("JZSms.JZAPI") '必须先要注册JZSms.dll组件
obj.releaseAll
obj.setWebService "http://www.jianzhou.sh.cn:8080/JianzhouSMSWSServer/services/BusinessService"
dxname=obj.sendBatchMessage("sdk_lys231896","974xl215",phone,name1&"你好啊!") '用户名密码用销售人员给的测试账号
Response.Write "<br>"
Response.Write("dxname:" & CStr(dxname)) 

%>

</body>
</html>