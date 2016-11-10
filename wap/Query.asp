<!--#Include file="Head.asp"-->
<!--#include file="info/MsgClass.asp"-->
	</div>
<style>
.tijiao{
    color: rgba(255,255,255,1);
    text-decoration: none;
    background-color: rgba(219,87,5,1);
    font-family: 'Yanone Kaffeesatz';
    font-weight: 700;
    font-size: 14px;
    display: block;
    padding: 4px;
    -webkit-border-radius: 8px;
    -moz-border-radius: 8px;
    border-radius: 8px;
    border: 0;
	width: 80px;
	text-align: center;
	-webkit-transition: all .1s ease;
	-moz-transition: all .1s ease;
	-ms-transition: all .1s ease;
	-o-transition: all .1s ease;
	transition: all .1s ease;
}
</style>
    <div id="main"> 
      <div class="topad1"><img src="images/io_tops.jpg" /></div>
      <div class="html">
         <div class="html1"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" valign="top">
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top" class="bk_xb1 bk_zb bk_yb"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="mag">
            <tr>
              <td class="lh_22">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="40" align="left" style="background:#26B30F;color:#fff;">【<strong>发货查询</strong>】</td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="107"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="40" class="text6" style="padding:5px;">发货状态查询（查询是否已经发货以及发货时间，必须正确输入订货人姓名才可查询）：</td>
              </tr>
              <tr>
                <td style="padding:5px;">
                 <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <form name="Search" id="Search" method="post" style="margin:0px;" action="Query.asp?Action=Search#S" onSubmit="return CheckValue3();">
                  <tr>
                    <td width="30%" height="30" class="text5">订货人姓名：</td>
                    <td width="40%"><input type="text" name="KeyWord" style="width:100px;height:26px;border:1px solid #999;margin:10px auto;" value=""/></td>
                    <td width="30%"><table border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="77" height="24" class="text7"><input type="submit" class="tijiao" value="点击查询" /></td>
                      </tr>
                    </table>
                    </td>
                  </tr>
                  </form>
                </table></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="40" align="left" class="text4" style="background:#26B30F;color:#fff;">【<strong>超过  六   天未收到货的朋友请在下面留言</strong>】</td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="107"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td align="left" class="text8" style="padding:5px;"><%=GetValues("NwebCn_About","Content",52)%></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <form name="Search_Msg2" id="Search_Msg2" action="QueryMsgSave.asp" onSubmit="return CheckMsg2();" style="margin:0px;" method="post">
              <tr>
                <td width="50%" height="25" align="left">您订货时提交的姓名：</td>
                <td width="50%" height="25" colspan="2" align="left"><input name="UserName" id="UserName" style="width:100px;height:26px;border:1px solid #999;margin:10px auto;" />
                  <font color="ff0000">*</font>请填写真实信息</td>
              </tr>
              <tr>
                <td width="50%" height="25" align="left">订货的大概时间：</td>
                <td width="50%" colspan="2" align="left"><input type="text" style="width:50px;height:26px;border:1px solid #999;margin:10px auto;" name="Year" id="Year" value="" /> 年 <input type="text" style="width:26px;height:26px;border:1px solid #999;margin:10px auto;" name="Month" id="Month" value="" /> 月 <input type="text" style="width:26px;height:26px;border:1px solid #999;margin:10px auto;" name="Day" id="Day" value="" /> 日 xx年xx月x日</td>
              </tr>
              <tr>
                <td width="50%" height="25" align="left">您的联系电话：</td>
                <td width="50%" height="25" align="left"><input type="text" style="width:100px;height:26px;border:1px solid #999;margin:10px auto;" name="TelPhone" id="TelPhone" /></td>
                <td width="40" align="left"><table width="40" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="40" height="21" align="center"><input type="submit" name="tijiao" id="tijiao" class="tijiao" value="提交" /></td>
                  </tr>
                </table></td>
              </tr>
              </form>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30"><hr size="1" color="#26B30F" /></td>
          </tr>
          <tr>
            <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" id="S">
                <tr>
                  <td height="40" align="left" style="padding:5px;"><%=GetValues("NwebCn_About","Content",53)%></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
		<td height="32" align="left">下面是反馈回来的快递信息：</td>
		</tr>
                    <tr>
                      <td align="left">
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <form name="Search_Phone" id="Search_Phone" action="Query.asp?Action=Search_Phone#S" method="post" style="margin:0px;" onSubmit="return Check();">
                        <tr>
                          <td width="30%" height="32" class="text10">电话号码：</td>
                          <td width="30%"><input type="text" name="TelPhone" id="TelPhones" style="width:100%;height:26px;border:1px solid #999;margin:10px auto;" value="<%=Trim(Request("TelPhone"))%>" /></td>
                          <td width="20%" align="center">&nbsp;<font color="ff0000">*</font>必填</td>
                          <td width="20%"><table border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="77" height="24" class="text7"><input type="submit" class="tijiao" value="点击查询" /></td>
                              </tr>
                          </table></td>
                        </tr>
                        </form>
                      </table></td>
                    </tr>
                  </table></td>
                </tr>
            </table>
            </td>
          </tr>
        </table>
         <%
		 	Dim Action
			Action=Trim(Request("Action"))
			Select Case Action
				Case "Search":
					SearchKeyList
				Case "Search_Phone":
					ViewText	
				Case Else%>
				<div class="search_tit">最新发货通知</div><div style="height:auto; width:100%; line-height:26px; color:#030303; text-align:left;">
						<%
                            DIm Object
                            Set Object=New ViewClass
                            Object.Set_Page_Size(50)
                            Object.ViewList
                        %>
               </div>						
			<%End Select
		 %>       
        </td>
      </tr>
    </table>
           	  </td>
          </tr>
            <tr>
              <td height="15">&nbsp;</td>
            </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
</table>
      </div>
    </div>
  </div><!--#Include file="Foot.asp"-->
<script language="javascript">
<!--
function CheckMsg2()
{
	var UserName,TelPhone,Year,Month,Day,ShiJian;
	UserName=document.getElementById("UserName");
	Year=document.getElementById("Year");
	Month=document.getElementById("Month");
	Day=document.getElementById("Day");
	TelPhone=document.getElementById("TelPhone");
	if(UserName.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写姓名信息！");
		UserName.focus();
		return false;
	}
	if(Year.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填年份信息！");
		Year.focus()
		return false;
	}
	else
	{
		if((Year.value).search("^-?\\d+(\\.\\d+)?$")!=0)
		{
			alert("请输入正确的年份信息！");
			Year.select();
			return false;
		}
	}
	
	
	if(Month.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填月份信息！");
		Month.focus()
		return false;
	}
	else
	{
		if((Month.value).search("^-?\\d+(\\.\\d+)?$")!=0)
		{
			alert("请输入正确的月份信息！");
			Month.select();
			return false;
		}
	}
	
	if(Day.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填月份信息！");
		Day.focus()
		return false;
	}
	else
	{
		if((Day.value).search("^-?\\d+(\\.\\d+)?$")!=0)
		{
			alert("请输入正确的月份信息！");
			Day.select();
			return false;
		}
	}
	
	
	
	ShiJian=Year.value+"-"+Month.value+"-"+Day.value
	if(ShiJian.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写时间信息！");
		return false;
	}
	else
	{
		var patrn=/^((\d{2}(([02468][048])|([13579][26]))[\-\/\s]?((((0?[13578])|(1[02]))[\-\/\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\-\/\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\-\/\s]?((0?[1-9])|([1-2][0-9])))))|(\d{2}(([02468][1235679])|([13579][01345789]))[\-\/\s]?((((0?[13578])|(1[02]))[\-\/\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\-\/\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\-\/\s]?((0?[1-9])|(1[0-9])|(2[0-8]))))))(\s(((0?[0-9])|([1-2][0-3]))\:([0-5]?[0-9])((\s)|(\:([0-5]?[0-9])))))?$/;
		if (!patrn.exec(ShiJian))
		{
             alert("请填写正确的时间格式：yy-mm-d!");
			 return false;
        }

	}
	
	if(TelPhone.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写联系信息！");
		TelPhone.focus();
		return false;
	}
	else
	{
		if((TelPhone.value).search("^-?\\d+$")!=0)
		{
			alert("请填写正确的联系信息！");
			TelPhone.select();
			return false;
		}
	}
	return true;
}

function Check()
{
	var TelPhone=document.getElementById("TelPhones");
	if(TelPhone.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写电话码！");
		TelPhone.focus();
		return false;
	}
	else
	{
		 var pattern =/^[+]{0,1}(\d){1,3}[ ]?([-]?((\d)|[ ]){1,12})+$/;
		  if(!pattern.exec(TelPhone.value))
             {
              	alert("请输入正确的电话号码！");
				TelPhone.select();
				return false;
             }
	}
	return true;
}

function CheckValue3()
{
	var KeyWord=document.getElementById("KeyWord");
	if(KeyWord.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写订货人姓名！");
		KeyWord.focus();
		return false;
	}
	return true;
}
-->
</script>

<!--代码-->
<%Sub SearchKeyList()%>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#ffffff" style="margin-top:10px;">
        <tr>
            <td width="19%" height="28" align="center" bgcolor="#26B30F"><span style="font-weight: bold;color:#FEFFFD;">订 单 编 号</span></td>
            <td width="20%" align="center" bgcolor="#26B30F"><span style="font-weight: bold; color:#FEFFFD;">订货人姓名</span></td>
            <td width="22%" height="28" align="center" bgcolor="#26B30F"><span style="font-weight: bold;color:#FEFFFD;">下 单 时 间</span></td>
            <td width="26%" height="28" align="center" bgcolor="#26B30F"><span style="font-weight: bold;color:#FEFFFD;">定 单 状 态</span></td>
            <td width="13%" align="center" bgcolor="#26B30F"><span style="font-weight: bold;color:#FEFFFD;">定 单 金 额</span></td>
        </tr>
   			<%Call SearchList(20)%>
    </table>
<%End Sub%>

<%
Sub SearchList(Page_Size)
	Dim KeyWord
	KeyWord=Trim(safeRequest("KeyWord","auto"))
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	if KeyWord<>"" then
		sql="select id,ProductNo,Linkman,AddTime,State,HuoDao_FuKuan,Remark,FuKuan,FaHuoTime from NwebCn_Order where Linkman ='"&KeyWord&"' order by AddTime desc"
	else
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('请输入查寻用户名！');"&vbcrlf)
			response.Write("window.history.go(-1);")
		response.Write("</script>")
		response.End()
		exit sub
	end if
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<tr bgcolor='#EBF6FC' height='25'>")
			response.Write("<td colspan='5' align='left' style='padding:5px;'>")
				response.Write("对不起，暂没有找你要的信息！")
			response.Write("</td>")
		response.Write("</tr>")
	else
		rs.pagesize=Page_Size
		dim sum_page,total,i
		total=rs.recordcount
		sum_page=total \ page_size
		if total mod page_size <>0 then sum_page=sum_page+1
		dim page
		page=trim(saferequest("page","get"))
		if page="" or isnull(page) or (not IsNumeric(page)) then
			page=1
		elseif Cint(Page)<=1 then
			page=1
		elseif Cint(page) => sum_page then
			page=sum_page
		else
			page=Cint(page)
		end if
		rs.absolutepage=page
		dim Flage
		 
		for i=1 to Page_Size
			if not rs.eof then
			flage=1
				response.Write("<tr bgcolor='#EBF6FC' height='25'>")
					response.Write("<td style='padding-left:5px;'>")
						'response.Write("<a href='OrderView.asp?ID="&rs("id")&"' target='_blank'>")
						response.Write(rs("ProductNo"))
						'response.Write("</a>")
					response.Write("</td>")
					response.Write("<td style='padding-left:5px;'>")
						response.Write(rs("Linkman"))
					response.Write("</td>")
					response.Write("<td style='padding-left:5px;'>")
						response.Write(rs("AddTime"))
					response.Write("</td>")
					response.Write("<td align='left' style='padding-left:5px;'>")
						if rs("State")<>"" then 
							response.Write(rs("State"))
						else
							response.Write("待处理……")						
						end if

						if flage=1 then
						end if
					'if rs("HuoDao_FuKuan") then
						'if rs("FuKuan") then
							'if rs("State")="货到后付款" then
								'response.Write("等待发货……")
							'else
								'if Instr(rs("State"),"货已发")>0 then
									'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
									'response.Write("&nbsp;&nbsp;发货时间："&FormatDate(rs("FaHuoTime"),4))
								'else
									'response.Write(rs("State"))
								'end if
							'end if
						'else
							'if rs("State")="" or isnull(rs("State")) then
								'response.Write("待处理……")
							'else
							'response.Write("对不起，当地不能货到付款，货没有发！")
							'end if
						'end if
					'else
						'if Instr(rs("State"),"货已发")>0 then
							'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
							'response.Write("&nbsp;&nbsp;发货时间："&FormatDate(rs("FaHuoTime"),4))
						'else
						'Response.Write("因你的收货地不支持货到付款，请重下订单先汇款才可以发货，如果有问题请电话咨询我们。 ")
							'response.Write(rs("State"))
						'end if
					'end if
					
					response.Write("</td>")
					
					response.Write("<td style='padding-left:5px;'>")
						response.Write(SumMemony(rs("Remark")))
					response.Write("</td>")
				response.Write("</tr>")
				rs.movenext
			end if
		next
		if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
	end if
End sub

sub Contrl_Page(page,sum_page,total,page_size) 
	dim Url,linkfile,pagewhere,UrlValue
	Url=request.ServerVariables("URL")
	Url=mid(Url,InstrRev(Url,"/")+1)
	linkfile=Url
	
	if trim(Request("KeyWord"))<>"" then
		Pagewhere="&KeyWord="&trim(safeRequest("KeyWord","auto"))
	end if
	
	if trim(Request("Action"))<> "" then
		Pagewhere="&Action="&trim(Request("Action"))
	end if
	
		response.Write("<tr>")
			response.Write("<td colspan=5 align='right' style='padding-top:5px;padding-bottom:5px;padding-right:10px;'>")
				response.Write("[共计："&total&"条] ")
						response.write("[每页："&page_size&"条] ")
						response.write("[页次："&page&"/"&sum_page&"] ")
						if page<=1 then
							response.write("[首页]　[上一页] ")
						else 
							response.write("[<a href='"&linkfile&"?page=1"&pagewhere&"'>")
							response.write("首页")
							response.write("</a>] ")
							response.write("[<a href='"&linkfile&"?page="&page-1&pagewhere&"'>")
							response.write("上一页")
							response.write("</a>] ")
						end if
						
						if page < sum_page then
							response.write("[<a href='"&linkfile&"?page="&page+1&pagewhere&"'>")
							response.write("下一页")
							response.write("</a>]　")
						else
							response.write("[下一页] ")
						end if
						
						if sum_page>1 and page < sum_page then
							response.write("[<a href='"&linkfile&"?page="&sum_page&pagewhere&"'>")
							response.write("末页")
							response.write("</a>]")
						else
							response.write("[末页]")
						end if
						dim cc
						response.write(" 转到：")%>
						<select name="page" size="1" onChange="javascript:window.location='<%=linkfile%>?page='+this.options[this.selectedIndex].value+'<%=pagewhere%>';">
							<%for cc=1 to sum_page
								if cc=page then
									response.write("<option value='"&cc&"' selected >"&cc&"页")
								else
									response.write("<option value='"&cc&"'>"&cc&"页")
								end if
							next%>
						</select>
			<%response.Write("</td>")
			response.Write("</tr>")
	end sub
	
	Function SumMemony(Remark)
		Dim Str
		Str=Split(Remark,"|")
		SumMemony=Str(ubound(Str))
	End function
	
	
	Sub ViewText()
		Dim TelPhone,Object
		TelPhone=Trim(safeRequest("TelPhone","auto"))
		if TelPhone="" or isnull(TelPhone) or Not(IsNumeric(TelPhone)) then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('数据出错，请返回！');"&vbcrlf)
				response.Write("window.history.go(-1);")
			response.Write("</script>"&vbcrlf)
			response.End()
		end if
		
		Set Object=New ViewClass
		Object.Set_TelNumber(TelPhone)
		Object.Set_Page_Size(7)
		Response.Write("<div ID='Msg_Border'>")
            Response.Write("<div id='WritMsg_Title'>")
                Response.Write("<strong>留言信息查看</strong>")
            Response.Write("</div>")
            Response.Write("<div id='MsgWrite_Text'>")
              Response.Write("<div style='text-align:center; height:auto;' id='MsgView_Text'>")
               	if Object.IsTrue then
					Object.ViewContent
                else
                    Response.Write("<div style='border:#B6DAEA 1px solid; background:#FFFFFF; height:120px; width:90%; line-height:20px; color:#030303; padding:10px; text-align:left;'>")
                    		Response.Write("对不起，你尚无留言信息！")
                        Response.Write("</div>")
                 end if
           	  Response.Write("</div>")
            Response.Write("</div>")
        Response.Write("</div>")
	End Sub
%>