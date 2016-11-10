<style type="text/css">
<!--
.STYLE6 {	font-size: 18;
	color: #FF0000;
}
.STYLE7 {
	font-family: "黑体";
	font-size: 18px;
}
-->
</style>
<div class="listRight" style="height:auto; width:100%; margin:0 auto;">
  <div class="div3 pd6" style="height:auto; width:96%; margin:0px; margin-top:10px;">
		  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin:auto; text-align:left; margin-top:10px;margin-bottom:13px;">
           	<form name="On_Order" id="On_Order" method="post" action="AliPay.asp?Action=ArrayPlay" onsubmit="return Check_OnOrder();">
<%
function getipadd()
 ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
 if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR")
 getipadd=ipadd
end function
%>
	<input type="hidden" name="ipadd" value="<%=getipadd%>">            
            <input type="hidden"  name="On_dgtime" id="On_dgtime" value="<%=Now()%>">
              <%
			Dim OrderId,str
			Dim THISO:THISO=str&right(year(now),1)&month(now)&day(now)&XXL(5)
			OrderId = HaveOrderId(str,THISO)
			%>
              <input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>">
              <input name="On_ZipCode" type="hidden" class="input4" id="On_ZipCode" size="10" maxlength="6" value="000000" />
              <input type="hidden" name="HuiKuan" value="支付宝" />
            <tr>
              <td height="30" colspan="4" style="line-height:25px;">
              	<%Call ProdList2()%>
              </td>
            </tr>
            <tr>
              <td width="30%" height="30">收 货 人：</td>
              <td width="70%" height="30" colspan="3"><input name="On_ShName" type="text" class="input4" size="15" id="On_ShName" /> <font color="#ff0000">*</font></td>
            </tr>
            <tr>
              <td width="30%" height="30">手机号码：</td>
              <td width="70%" height="30" colspan="3"><input name="On_ShTel" type="text" class="input4" size="15" id="On_ShTel" /> <font color="#ff0000">*</font></td>
            </tr>
            <tr>
              <td width="30%" height="30">收货地址：</td>
              <td width="70%" height="30" colspan="3"><input name="On_Addres" type="text" class="input4" style="width:92%" id="On_Addres" /> <font color="#ff0000">*</font></td>
            </tr>
            <tr>
              <td height="30" colspan="4" align="center" valign="bottom"><input id="alipay" type="image" src="../../images/btn05.jpg" /></td>
            </tr>
            </form>
          </table>
  </div>
	  </div>
      <%
  	  sub ProdList2()
    	dim rs,sql
        set rs=server.CreateObject("adodb.recordset")
		sql="select ProductName,Price,Price2,PriceText from NwebCn_Products where ViewFlag=1 order by px asc"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("暂无产品信息，不能定购！")
			response.Write("<input type='hidden' name='On_RecordCount' id='On_RecordCount' value='0'>")
		else
			dim i
			i=1
    		while not rs.eof 
				response.Write("<strong>"&rs("ProductName")&"</strong>&nbsp;")
				response.Write("<font color='#ff0000'>支付宝付款"&rs("Price2")&rs("PriceText")&"</font>")
				response.Write("<label>")
					response.Write("<select name='On_Numbers"&i&"' size='1' id='On_Numbers"&i&"'>")
						
						response.Write("<option value='NULL' selected>选择订购数量</option>")
						response.Write("<option value='"&rs("ProductName")&"(0)'>0盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(1)'>1盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(2)'>2盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(3)'>3盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(4)'>4盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(5)'>5盒</option>")
					response.Write("</select>")
				response.Write("</label>")
				response.Write("<br />")
				rs.movenext
				i=i+1
			wend
			response.Write("<input type='hidden' name='On_RecordCount' id='On_RecordCount' value='"&rs.recordcount&"'>")
		end if
		rs.close()
		set rs=Nothing
    End sub
    %>
	
    <script language="javascript">
	<!--
	function Check_OnOrder()
	{
		var On_dgtime,On_ShName,On_ShMoble,On_ShTel,On_Addres,HuiKuan,On_RecordCount,On_ZipCode
		On_dgtime=document.getElementById("On_dgtime");
		On_ShName=document.getElementById("On_ShName");
		On_ShTel=document.getElementById("On_ShTel");
		On_Addres=document.getElementById("On_Addres");
		On_ZipCode=document.getElementById("On_ZipCode");
		HuiKuan=document.getElementById("HuiKuan");
		On_RecordCount=document.getElementById("On_RecordCount");
		
		if(On_dgtime.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("数据出错，请刷新网页！");
			return false;
		}
		if(On_ShName.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写收货人信息！");
			On_ShName.focus();
			return false;
		}
		
		/*if(On_ShMoble.value.replace(/^\s*|\s*$/g,'')!="")
		{
			var moble=On_ShMoble.value;
			var patrn1=/^[+]{0,1}(\d){1,3}[ ]?([-]?((\d)|[ ]){1,12})+$/;
			if(!patrn1.exec(moble))
			{
				alert("请填写正确的电话号码！");
				On_ShMoble.select();
				return false;
			}
		}
		*/
		
		if(On_ShTel.value.replace(/^\s*|\s*$/g,'')!="")
		{
		//	var tel=On_ShTel.value
		//	var patrn2=/^[+]{0,1}(\d){1,4}[ ]?([-]?((\d)|[ ]){1,12})+$/;
		//	if(!patrn2.exec(tel))
		//	{
		//		alert("请填写收货人联系信息！");
		//		On_ShTel.focus();
		//		return false;
		//	}
		}
		
		
		if(On_ShTel.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写一个用户联系方式！");
			On_ShTel.focus();
			return false;
		}
		if(On_Addres.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写收货人地址！");
			On_Addres.focus();
			return false;
		}
		if(On_ZipCode.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写邮编！");
			On_ZipCode.focus();
			return false;
		}
		if(HuiKuan.value=="NULL")
		{
			alert("请先选择付款方式！");
			return false;
		}
		if(parseInt(On_RecordCount.value)<=0)
		{
			alert("暂无商品可以定购，不能断续！");
			return false;
		}
		else
		{
			var falg=false;
			for(var i=1;i<=parseInt(On_RecordCount.value);i++)
			{
				if(document.getElementById("On_Numbers"+i).value!="NULL")
				{
					falg=true;
				}
			}
			if(!falg)
			{
				alert("请选择定购商品的数量！");
				return false;
			}
		
		}
		return true;
	}
	
	-->
	</script>