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
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>查看、修改、回复订单</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|81,")=0 then 
  response.write ("<script language=javascript> alert('你不具有该管理模块的操作权限，请返回！');history.back(-1);</script>")
end if
%>
<%
if Instr(session("AdminPurview"),"|305,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ReplyContent,ReplyTime,ID,ProductName,ProductNo,Amount,Remark,display,NotSend
dim Linkman,Company,Address,ZipCode,Telephone,Fax,Mobile,Email,AddTime,States,FuKuan,HuoDao_FuKuan,Tel
ID=request.QueryString("ID")
call OrderEdit() 
display="none"
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>订单信息：查看，修改，回复订单信息相关的内容</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="OrderList.asp" onClick='changeAdminFlag("订单信息列表")'>查看订单信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>网站信息设置</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="OrderEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">商品名称：</td>
        <td><input name="ProductName" type="text" class="textfield" id="ProductName" style="WIDTH: 240;" value="<%=ProductName%>">&nbsp;&nbsp;
        	<%if HuoDao_FuKuan then%>
            <input type="hidden" name="HuoDao_FuKuan" id="HuoDao_FuKuan" value="1">
        	<input name="FuKuan" id="FuKuan" type="checkbox" value="1" <%if FuKuan then response.Write("Checked")%>>&nbsp;货到后付款
            <%else%>
            <input type="hidden" name="HuoDao_FuKuan" id="HuoDao_FuKuan" value="0">
        	<%end if%>
        </td>
      </tr>
      <tr>
        <td height="20" align="right">商品编号：</td>
        <td><input name="ProductNo" type="text" class="textfield" id="ProductNo" style="WIDTH: 240;" value="<%=ProductNo%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">订购数量：</td>
        <td><input name="Amount" type="text" class="textfield" id="Amount"   value="<%=Amount%>" size="80"><%=print(Amount)%></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">补充说明：
        <td><textarea name="Remark" rows="6" class="textfield" id="Remark" style="WIDTH: 76%;"><%=PringText(Remark)%></textarea></td>
      </tr>
      <tr>
        <td height="20" align="right">订&nbsp;购&nbsp;者：</td>
        <td><%=Linkman%></td>
      </tr>
      <tr>
        <td height="20" align="right">单位名称：</td>
        <td><input name="Company" type="text" class="textfield" id="Company" style="WIDTH: 240;" value="<%=Company%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">通信地址：</td>
        <td><input name="Address" type="text" class="textfield" id="Address"  value="<%=Address%>" size="80"></td>
      </tr>
      <tr>
        <td height="20" align="right">区　　号：</td>
        <td><input name="ZipCode" type="text" class="textfield" id="ZipCode" style="WIDTH: 120" value="<%=ZipCode%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">电　　话：</td>
        <td><input name="Telephone" type="text" class="textfield" id="Telephone" style="WIDTH: 240;" value="<%=Tel%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">传　　真：</td>
        <td><input name="Fax" type="text" class="textfield" id="Fax" style="WIDTH: 120" value="<%=Fax%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">移动电话：</td>
        <td><input name="Mobile" type="text" class="textfield" id="Mobile" style="WIDTH: 120" value="<%=Mobile%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">电子邮箱：</td>
        <td><input name="Email" type="text" class="textfield" id="Email" style="WIDTH: 240" value="<%=Email%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">订购时间：</td>
        <td><input name="AddTime" type="text" class="textfield" id="AddTime" style="WIDTH: 240" value="<%=AddTime%>"></td>
      </tr>
      <tr>
        <td height="30" align="right">修改定单状态：          </td>
        <td valign="bottom"><label>
        <select name="Stats" id="Stats" onChange="Event_Chang()">
			<%if HuoDao_FuKuan and (instr(States,"货已发")<=0 or States="" or isnull(States)) then%>
            	<option value="待处理">待处理</option>
            	<%if instr(States,"货到付款")>0 then%>
                <option value="货到付款" selected>货到付款</option>
                <%else%>
                <option value="货到付款">货到付款</option>
                <%end if%>
                <%if instr(States,"不能发货")>0 then%>
                <option value="不能发货" selected>不能发货</option>
				<%display=""%>
                <%else%>
                <option value="不能发货">不能发货</option>
                <%end if%>
            <%else%>        
				<%if instr(States,"未付款")>0 then%>
                  <option value="未付款" selected>未付款</option>
                <%else%>
                  <option value="未付款">未付款</option>
                <%end if%>
                <%if instr(States,"未发货")>0 then%>
                	<option value="未发货" selected>未发货</option>
                     <%display=""%>
                <%else%>
                	<option value="未发货">未发货</option>
                <%end if%>
                <%if instr(States,"已付款")>0 then%>
                  <option value="已付款" selected>已付款</option>
                <%else%>
                  <option value="已付款">已付款</option>
                <%end if%>
                
                <%if instr(States,"货已发")>0 then%>
                  <option value="货已发" selected>货已发</option>
                <%else%>
                    <option value="货已发">货已发</option>
                <%end if%>
                
                <%if instr(States,"货未收到")>0 then%>
                  <option value="货未收到" selected>货未收到</option>
                <%else%>
                    <option value="货未收到">货未收到</option>
                <%end if%>
            <%end if%>
         </select>
         <script language="javascript">
		 	<!--
			
			
			function Event_Chang()
			{
				var Stats,NotSend;
				Stats=document.getElementById("Stats");
				NotSend=document.getElementById("NotSend");
				if((Stats.value).indexOf("未发货")!=-1 || (Stats.value).indexOf("不能发货")!=-1)
				{
					NotSend.style.display="";
				}
				else
				{
					NotSend.style.display="none";
				}
			}
			-->
		 </script>
         <span style="margin-left:20px;display:<%=display%>;" id="NotSend">
         	<input type="text" name="NotSend" size="50" value="<%=NotSend%>"/>&nbsp;<font color="#FF0000">*请填写原因</font>
         </span>
        </label></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom"><label>
          <input type="submit" name="Modify" id="Modify" value="修 改">
          <input type="button" name="Modify2" id="Modify2" value="返 回" onClick="window.history.go(-1);">
        </label></td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub OrderEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_Order where ID="&ID
      rs.open sql,conn,1,3
	  if trim(Request.Form("HuoDao_FuKuan"))="1" then
		  if Trim(Request.Form("FuKuan"))="1" then
			  rs("FuKuan")=true
			  rs("State")=StrReplace(Request.Form("Stats"))
		  else
			  rs("FuKuan")=false
		  end if
	  else
	  	rs("State")=StrReplace(Request.Form("Stats"))	  
	  end if
	  Rs("Amount")=Trim(Request.Form("Amount"))
	 ' response.Write(Replace(Replace(Replace(Replace(Trim(Request.Form("Remark")),"支付方式：","|"),"应付金额：","|"),"送货方式：",""),vbcrlf,""))
	 ' Response.End()
	  Rs("Remark")=Replace(Replace(Replace(Replace(Trim(Request.Form("Remark")),"支付方式：","|"),"应付金额：","|"),"送货方式：",""),vbcrlf,"")
	  rs("Company")=Trim(Request.Form("Company"))
	  rs("Address")=Trim(Request.Form("Address"))
	  rs("ZipCode")=Trim(Request.Form("ZipCode"))
	  rs("Tel")=Trim(Request.Form("Telephone"))
	  rs("Fax")=Trim(Request.Form("Fax"))
	  rs("Telephone")=Trim(Request.Form("Mobile"))
	  rs("Email")=Trim(Request.Form("Email"))
	  rs("AddTime")=Trim(Request.Form("AddTime"))
	  if Trim(Request.Form("NotSend"))<>"" then
	  	rs("NotSendText")=trim(Request.Form("NotSend"))
	  end if
	  if instr(Trim(Request.Form("Stats")),"货已发")>0 then
	  	rs("FaHuoTime")=Now()
	  end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑订单信息！');changeAdminFlag('订单信息列表');location.replace('OrderList.asp');</script>"
  else '提取留言信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Order where ID="& ID
      rs.open sql,conn,1,1
	  ProductName=rs("ProductName")
	  ProductNo=rs("ProductNo")
	  Amount=rs("Amount")
	  Remark=ReStrReplace(rs("Remark"))
	  Linkman=GuestInfo(rs("MemID"),rs("Linkman"),rs("Sex"))
	  Company=rs("Company")
	  Address=rs("Address")
	  ZipCode=rs("ZipCode")
	  FuKuan=rs("FuKuan")
	  States=rs("State")
	  NotSend=rs("NotSendText")
	  Tel=rs("Tel")
	  Fax=rs("Fax")
	  Mobile=rs("Telephone")
	  Email=rs("Email")
	  HuoDao_FuKuan=rs("HuoDao_FuKuan")
	  AddTime=rs("AddTime")
	  ReplyContent=ReStrReplace(rs("ReplyContent"))
	  ReplyTime=rs("ReplyTime")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub

function GuestInfo(ID,Guest,Sex)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    GuestInfo=Guest & "&nbsp;" & Sex
  else
    GuestInfo="<font color='green'>会员&nbsp;</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"' onClick='changeAdminFlag(""前台会员资料"")'>"&Guest&"</a>"&Sex
  end if
  rs.close
  set rs=nothing
end function 

function Print(Amount)
	dim str,i,str1,str2,str3
	str1=""
	str=split(Amount,"|")
	for i=0 to ubound(str)
		if i>0 then str1=str1&"、"
		if str1="" then
			str1=Mid(str(i),1,instr(str(i),"(")-1)
			str2=Mid(str(0),instr(str(i),"(")+1,1)
			str3=""
		else
			str1=str1&Mid(str(i),1,instr(str(i),"(")-1)
			str2=Mid(str(0),instr(str(i),"(")+1,1)
			str3=Mid(str(1),instr(str(i),"(")+1,1)
		end if
		str1=str1&Mid(str(i),instr(str(i),"(")+1,(instr(str(i),")"))-(instr(str(i),"(")+1))&"盒"
	next
	Print=str1&"||"&str2&"||"&str3
end function

function PringText(Remark)
	dim str,str1,i
	str=split(Remark,"|")
	if ubound(str)>0 then
	str1="送货方式："&str(0)
	str1=str1&vbcrlf
	str1=str1&"支付方式："&str(1)
	str1=str1&vbcrlf
	str1=str1&"应付金额："&str(2)
	PringText=str1
	end if
end function
%>