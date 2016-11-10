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
<TITLE>网站信息设置</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<%
call CreateEditor("gonggao")
%>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|112,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<body>
<%
dim ID,SiteTitle,SiteUrl,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,IcpNumber,SystemSN,syimg,gonggao,QQ,syjs,qq2,taobaoid,otherscount,jobcount,OrderSates
dim MesViewFlag,zfbKey,zfbid,WY_ID,WY_Key
Dim smsID1,smsPWD1,smsID2,smsPWD2,MSG1,MSG2,MSG3,MSG4,MSG5
Dim leftGonggaoTitle,leftGonggaoContent,leftGonggaoView,leftGonggaoWidth,message_note

select case request.QueryString("Action")
  case "Save"
    SaveSiteInfo
  case else
    ViewSiteInfo
end select
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>系统管理：添加，修改站点的相关信息</strong></font></td>
  </tr>
 <!-- <tr>
    <td height="24" align="center" nowrap bgcolor="#EBF2F9">
	<a href="PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("修改密码")'>修改密码</a>	<font color="#0000FF">&nbsp;|&nbsp;</font>	<a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>网站信息设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NavigationList.asp" target="mainFrame" onClick='changeAdminFlag("栏目导航设置")'>栏目导航设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetConst.asp" target="mainFrame" onClick='changeAdminFlag("常量设置")'>常量设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="DataManage.asp" target="mainFrame" onClick='changeAdminFlag("数据库操作")'>数据库操作</a>
<font color="#0000FF">&nbsp;|&nbsp;</font><a href="ADsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("弹窗广告列表")'>弹窗广告</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SpaceStat.asp" target="mainFrame" onClick='changeAdminFlag("空间统计")'>空间统计</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="../Count/InfoList.asp" target="mainFrame" onClick='changeAdminFlag("访问统计")'>访问统计</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="FriendSiteList.asp" target="mainFrame" onClick='changeAdminFlag("友情链接")'>友情链接</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="HackSql.asp" target="mainFrame" onClick='changeAdminFlag("阻止SQL注入记录")'>阻止SQL注入记录</a>    </td>
  </tr>-->
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="SetSite.asp?Action=Save" >
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">网站标题：</td>
        <td><input name="SiteTitle" type="text" class="textfield" id="SiteTitle" style="WIDTH: 400;" value="<%=SiteTitle%>" >&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">网　　址：</td>
        <td><input name="SiteUrl" type="text" class="textfield" id="SiteUrl" style="WIDTH: 400;" value="<%=SiteUrl%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">公司名称：</td>
        <td><input name="ComName" type="text" class="textfield" id="ComName" style="WIDTH: 400;" value="<%=ComName%>" >&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">地　　址：</td>
        <td><input name="Address" type="text" class="textfield" id="Address" style="WIDTH: 400;" value="<%=Address%>" >&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">邮　　编：</td>
        <td><input name="ZipCode" type="text" class="textfield" id="ZipCode" style="WIDTH: 200;" value="<%=ZipCode%>" maxlength="20">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">电　　话：</td>
        <td><input name="Telephone" type="text" class="textfield" id="Telephone" style="WIDTH: 200;" value="<%=Telephone%>">
        &nbsp;* <span class="STYLE1">填写2个电话请用&quot;|&quot;隔开 要不不能识别</span> </td>
      </tr>
      <tr>
        <td height="20" align="right">传　　真：</td>
        <td><input name="Fax" type="text" class="textfield" id="Fax" style="WIDTH: 200;" value="<%=Fax%>" >&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">电子邮箱：</td>
        <td><input name="Email" type="text" class="textfield" id="Email" style="WIDTH: 200;" value="<%=Email%>">&nbsp;*&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">QQ：</td>
        <td><input name="QQ" type="text" class="textfield" id="QQ" style="WIDTH: 200;" value="<%=qq%>">&nbsp;</td>
      </tr>
	  <tr>
        <td height="20" align="right">QQ2：</td>
        <td><input name="QQ2" type="text" class="textfield" id="QQ2" style="WIDTH: 200;" value="<%=qq2%>">&nbsp;</td>
      </tr>
	     <tr>
	       <td height="20" align="right">支付宝帐号：</td>
	       <td><input name="zfbid" type="text" class="textfield" id="zfbid" style="WIDTH: 200;" value="<%=zfbid%>"></td>
          </tr>
          
	     <tr>
	       <td height="20" align="right">短信KEY：</td>
	       <td><input name="smsID1" type="text" class="textfield" id="smsID1" style="WIDTH: 200;" value="<%=smsID1%>"><span id="smsId1_count"></span></td>
          </tr>
	     <tr>
	       <td height="20" align="right">短信PWD：</td>
	       <td><input name="smsPWD1" type="text" class="textfield" id="smsPWD1" style="WIDTH: 200;" value="<%=smsPWD1%>"></td>
          </tr>
	     <tr style="display:none">
	       <td height="20" align="right">货到付款短信KEY：</td>
	       <td><input name="smsID2" type="text" class="textfield" id="smsID2" style="WIDTH: 200;" value="<%=smsID2%>"><span id="smsId2_count"></span></td>
          </tr>
          

	     <tr  >
	       <td height="20" align="right">【货到付款】短信：</td>
	       <td>
           <textarea name="MSG1" rows="2"  class="textfield" id="MSG1" style="WIDTH: 400;"><%=MSG1%></textarea> √ 一条短信70个字
           </td>
         </tr>
	     <tr  >
	       <td height="20" align="right">【银行和支付宝】短信：</td>
	       <td>
           <textarea name="MSG3" rows="2"  class="textfield" id="MSG3" style="WIDTH: 400;"><%=MSG3%></textarea> 一条短信70个字
           </td>
         </tr>
	     <tr >
	       <td height="20" align="right">【支付宝】短信：</td>
	       <td>
           <textarea name="MSG5" rows="2"  class="textfield" id="MSG5" style="WIDTH: 400;"><%=MSG5%></textarea> 一条短信70个字
           </td>
         </tr>
	     <tr  >
	       <td height="20" align="right">【支付宝成功】短信：</td>
	       <td>
           <textarea name="MSG4" rows="2"  class="textfield" id="MSG4" style="WIDTH: 400;"><%=MSG4%></textarea> 一条短信70个字
           </td>
         </tr>
	     <tr>
	       <td height="20" align="right">【发货】短信：</td>
	       <td>
           <textarea name="MSG2" rows="2"  class="textfield" id="MSG2" style="WIDTH: 400;"><%=MSG2%></textarea> √ 一条短信70个字
           </td>
          </tr>
          
	     <tr>
	       <td height="20" align="right">安全校验码：</td>
	       <td><input name="zfbKey" type="text" class="textfield" id="zfbKey" style="WIDTH: 200;" value="<%=zfbKey%>"></td>
          </tr>
	     <tr>
	       <td height="20" align="right">网银商户号：</td>
	       <td><input name="WY_ID" type="text" class="textfield" id="WY_ID" style="WIDTH: 200;" value="<%=WY_ID%>"></td>
          </tr>
	     <tr>
	       <td height="20" align="right"> MD5私钥：</td>
	       <td><input name="WY_Key" type="text" class="textfield" id="WY_Key" style="WIDTH: 200;" value="<%=WY_Key%>"></td>
          </tr>
	  <tr>
        <td height="20" align="right">订单状态：</td>
        <td><input name="OrderSates" type="text" class="textfield" id="OrderSates" style="WIDTH:400;" value="<%=OrderSates%>">&nbsp;以|号间隔</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">关 键 字：</td>
        <td><textarea name="Keywords" rows="6"  class="textfield" id="Keywords" style="WIDTH: 400;"><%=Keywords%></textarea>&nbsp;关键字设置有利于网站的搜索</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">网站描述：</td>
        <td><textarea name="Descriptions" rows="6" class="textfield" id="Descriptions" style="WIDTH: 400;"><%=Descriptions%></textarea>&nbsp;网站描述设置有利于网站的搜索</td>
      </tr>
         <tr>
        <td height="20" align="right">网站公告标题：</td>
        <td><input name="taobaoid" type="text" class="textfield" id="taobaoid" style="WIDTH: 200;" value="<%=taobaoid%>">&nbsp;</td>
      </tr>
	   <tr>
        <td height="20" align="right" valign="top">网站公告：</td>
        <td><textarea name="gonggao" rows="12" class="textfield" id="gonggao" style="WIDTH: 400;"><%=gonggao%></textarea>&nbsp;网站公告</td>
      </tr>
      <tr>
        <td height="20" align="right">是否显示公告：</td>
        <td><input name="otherscount" type="checkbox" id="otherscount" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if otherscount then response.write ("checked")%>>&nbsp;是否显示公告</td>
      </tr>
      <tr style="display:none;">
        <td height="20" align="right">是否显示公告详细：</td>
        <td><input name="jobcount" type="checkbox" id="jobcount" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if jobcount then response.write ("checked")%>>&nbsp;是否显示公告详细</td>
      </tr>
	  <tr>
        <td height="20" align="right" valign="top">公告框宽高：</td>
        <td><input name="syjs" type="text" class="textfield" id="syjs" style="WIDTH: 200;" value="<%=syjs%>">&nbsp;“300|200”以|分割</td>
      </tr>
      
	    <tr style="display:none">
        <td height="20" align="right" valign="top">左边公告标题：</td>
        <td><input name="leftGonggaoTitle" type="text" class="textfield" id="leftGonggaoTitle" style="WIDTH: 200;" value="<%=leftGonggaoTitle%>"></td>
      </tr>
	   <tr style="display:none">
        <td height="20" align="right" valign="top">左边公告内容：</td>
        <td><textarea name="leftGonggaoContent" rows="12" class="textfield" id="leftGonggaoContent" style="WIDTH: 400;"><%=leftGonggaoContent%></textarea>&nbsp;</td>
      </tr>
	  <tr style="display:none">
        <td height="20" align="right" valign="top">左边公告框宽高：</td>
        <td><input name="leftGonggaoWidth" type="text" class="textfield" id="leftGonggaoWidth" style="WIDTH: 200;" value="<%=leftGonggaoWidth%>">&nbsp;“300|200”以|分割</td>
      </tr>
	  <tr style="display:none">
        <td height="20" align="right">是否显示左边公告：</td>
        <td><input name="leftGonggaoView" type="checkbox" id="leftGonggaoView" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if leftGonggaoView then response.write ("checked")%>>&nbsp;是否显示左边公告</td>
      </tr>
	  <tr>
        <td height="20" align="right" valign="top">留言后提示信息：</td>
        <td><textarea name="message_note" rows="3" class="textfield" id="message_note" style="WIDTH: 400;"><%=message_note%></textarea></td>
      </tr>
	  <tr>
  <td height="20" align="right">首页简介图：</td>
        <td><input name="syimg" type="text" class="textfield" style="WIDTH: 240;" value="<%=syimg%>" maxlength="100">
        &nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=syimg',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"> </a><span class="STYLE2">推荐162*135</span></td>
      </tr>
     <tr>
        <td height="20" align="right">ICP&nbsp;备案：</td>
        <td><input name="IcpNumber" type="text" class="textfield" id="IcpNumber" style="WIDTH: 200;" value="<%=IcpNumber%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">授&nbsp;权&nbsp;号：</td>
        <td><input name="SystemSN" type="text" class="textfield" id="SystemSN" style="WIDTH: 200;" value="<%=SystemSN%>" readonly></td>
      </tr>

      <tr>
        <td height="20" align="right"><a name="Message"></a>留&nbsp;言&nbsp;簿：</td>
        <td><input name="MesViewFlag" type="checkbox" id="MesViewFlag" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if MesViewFlag then response.write ("checked")%>>&nbsp;自动通过审核</td>
      </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 60;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</body>
</html>
<%
'conn.execute "ALTER TABLE NwebCn_Site ADD COLUMN OrderSates TEXT(255)"

function SaveSiteInfo()
  if len(trim(request.Form("SiteTitle")))<3 then
    response.write ("<script language=javascript> alert('网站标题必填且不能少于3个字符！');history.back(-1);</script>")
    response.end
  end if
   if len(trim(request.Form("SiteUrl")))<10 then
    response.write ("<script language=javascript> alert('网站网址不能为空，且不少于10个字符！');history.back(-1);</script>")
    response.end
  end if 
  if left(trim(request.Form("SiteUrl")),7)<>"http://" then
    response.write ("<script language=javascript> alert('网站网址请加上""http://""！');history.back(-1);</script>")
    response.end
  end if
  if len(trim(request.Form("ComName")))<3 then
    response.write ("<script language=javascript> alert('公司名称必填且不能少于3个字符！');history.back(-1);</script>")
    response.end
  end if 
  if len(trim(request.Form("Address")))<3 then
    response.write ("<script language=javascript> alert('公司地址必填且不能少于3个字符！');history.back(-1);</script>")
    response.end
  end if
  if len(trim(request.Form("ZipCode")))<6 then
    response.write ("<script language=javascript> alert('邮政编码必填且不能少于6个字符！');history.back(-1);</script>")
    response.end
  end if
  if len(trim(request.Form("Telephone")))<11 then
    response.write ("<script language=javascript> alert('电话号码必填且不能少于11个字符！');history.back(-1);</script>")
    response.end
  end if
  if len(trim(request.Form("Fax")))<11 then
    response.write ("<script language=javascript> alert('传真号码必填且不能少于11个字符！');history.back(-1);</script>")
    response.end
  end if
  if len(trim(request.Form("Email")))<6 then
    response.write ("<script language=javascript> alert('电子邮箱必填具不能少于6个字符！');history.back(-1);</script>")
    response.end
  end if
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from NwebCn_Site"
  rs.open sql,conn,1,3
  rs("zfbKey")=Trim(Request.Form("zfbKey"))
  rs("zfbid")=Trim(Request.Form("zfbid"))
  rs("WY_ID")=trim(Request.Form("WY_ID"))
  rs("WY_Key")=Trim(Request.Form("WY_Key"))
  rs("SiteTitle")=trim(Request.Form("SiteTitle"))
  rs("SiteUrl")=trim(Request.Form("SiteUrl"))
  rs("ComName")=trim(Request.Form("ComName"))
  rs("Address")=trim(Request.Form("Address"))
  rs("ZipCode")=trim(Request.Form("ZipCode"))
  rs("Telephone")=trim(Request.Form("Telephone"))
  rs("Fax")=trim(Request.Form("Fax"))
 Rs("syimg")=trim(Request.Form("Syimg"))
  rs("Email")=trim(Request.Form("Email"))
  rs("QQ")=trim(Request.Form("QQ"))
   rs("QQ2")=trim(Request.Form("QQ2"))
    rs("taobaoid")=trim(Request.Form("taobaoid"))
	rs("OrderSates")=request.Form("OrderSates")
  rs("Keywords")=trim(Request.Form("Keywords"))
  rs("Descriptions")=trim(Request.Form("Descriptions"))
    rs("gonggao")=trim(Request.Form("gonggao"))
	Rs("syjs")=Trim(Request.Form("Syjs"))
  rs("IcpNumber")=trim(Request.Form("IcpNumber"))
  rs("smsID1")=trim(Request.Form("smsID1"))
  rs("smsPWD1")=trim(Request.Form("smsPWD1"))
  rs("smsID1")=trim(Request.Form("smsID1"))
  rs("smsPWD2")=trim(Request.Form("smsPWD2"))
  rs("MSG1")=trim(Request.Form("MSG1"))
  rs("MSG2")=trim(Request.Form("MSG2"))
  rs("MSG3")=trim(Request.Form("MSG3"))
  rs("MSG4")=trim(Request.Form("MSG4"))
  rs("MSG5")=trim(Request.Form("MSG5"))
if Request.Form("otherscount")=1 then
    rs("otherscount")=1
  else
    rs("otherscount")=0
  end if
  if Request.Form("jobcount")=1 then
    rs("jobcount")=1
  else
    rs("jobcount")=0
  end if
  'rs("leftGonggaoTitle")=trim(Request.Form("leftGonggaoTitle"))
  'rs("leftGonggaoContent")=trim(Request.Form("leftGonggaoContent"))
  'rs("leftGonggaoView")=trim(Request.Form("leftGonggaoView"))
  'rs("leftGonggaoWidth")=trim(Request.Form("leftGonggaoWidth"))
  rs("message_note")=trim(Request.Form("message_note"))
  
  
  if Request.Form("MesViewFlag")=1 then
    rs("MesViewFlag")=Request.Form("MesViewFlag")
    'Conn.execute "ALTER TABLE NwebCn_Message ALTER COLUMN  ViewFlag bit default 1"
	'Conn.execute "ALTER TABLE NwebCn_Message ADD CONSTRAINT  [DF_NwebCn_Message_ViewFlag]   DEFAULT   (1)   FOR   [ViewFlag]"
  else
    rs("MesViewFlag")=0
    'Conn.execute "ALTER TABLE NwebCn_Message ALTER COLUMN  ViewFlag bit default 0"
  end if
  rs.update
  rs.close
  set rs=nothing 
  response.write "<script language=javascript> alert('成功编辑网站信息！');changeAdminFlag('网站信息设置');location.replace('SetSite.asp');</script>"
end function

function ViewSiteInfo()
  dim rs,sql 
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from NwebCn_Site"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write "读取数据库记录出错！"
    response.end
  else
    SiteTitle=rs("SiteTitle")
    SiteUrl=rs("SiteUrl")
    ComName=rs("ComName")
    Address=rs("Address")
	zfbKey=rs("zfbKey")
	zfbid=rs("zfbid")
    WY_ID=rs("WY_ID")
    WY_Key=rs("WY_Key")
    ZipCode=rs("ZipCode")
    Telephone=rs("Telephone")
    Fax=rs("Fax")
    Email=rs("Email")
	otherscount=Rs("otherscount")
	QQ=rs("QQ")
	qq2=rs("qq2")
	taobaoid=Rs("taobaoid")
	syimg=Rs("syimg")
	OrderSates=rs("OrderSates")
    Keywords=rs("Keywords")
    Descriptions=rs("Descriptions")
	gonggao=rs("gonggao")
    IcpNumber=rs("IcpNumber")
    SystemSN=rs("SystemSN")
	MesViewFlag=rs("MesViewFlag")
	syjs=Rs("syjs")
	jobcount=Rs("jobcount")
	smsID1=Rs("smsID1")
	smsID2=Rs("smsID2")
	smsPWD1=Rs("smsPWD1")
	smsPWD2=Rs("smsPWD2")
	MSG1=Rs("MSG1")
	MSG2=Rs("MSG2")
	MSG3=Rs("MSG3")
	MSG4=Rs("MSG4")
	MSG5=Rs("MSG5")
	
	'leftGonggaoTitle=Rs("leftGonggaoTitle")
	'leftGonggaoContent=Rs("leftGonggaoContent")
	'leftGonggaoView=Rs("leftGonggaoView")
	'leftGonggaoWidth=Rs("leftGonggaoWidth")
	
	message_note=Rs("message_note")
    rs.close
    set rs=nothing 
  end if
end function
%>
