
<%
Response.Charset="gb2312"
'�ж��Ƿ��¼���¼��ʱ===============================================================
dim AdminAction
AdminAction=request.QueryString("AdminAction")
select case AdminAction
  case "Out"
    call OutLogin()
  case else
    call Login()
end select
'========
sub Login()
if session("AdminName")="" then Session("AdminName")= Request.Cookies("AdminName")
if session("UserName")="" then Session("UserName")= Request.Cookies("UserName")
if session("AdminPurview")="" then Session("AdminPurview")= Request.Cookies("AdminPurview")
if session("LoginSystem")="" then Session("LoginSystem")= Request.Cookies("LoginSystem")
  if session("AdminName")="" or session("UserName")="" or session("AdminPurview")="" or session("LoginSystem")<>"Succeed" then
     response.write "����û�е�¼���¼�ѳ�ʱ����<a href='Login.asp' target='_parent'><font color='red'>���ص�¼</font></a>!"
     response.end
  end if
end sub
'========
sub OutLogin()
  session.contents.remove "AdminName"
  session.contents.remove "UserName"
  session.contents.remove "AdminPurview"
  session.contents.remove "LoginSystem"
  session.contents.remove "VerifyCode"
  response.write "<script language=javascript>top.location.replace('Login.asp');</script>"
end sub
%>