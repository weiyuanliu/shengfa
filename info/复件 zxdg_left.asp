<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE3 {color: #FF0000; font-size: 14; }
.STYLE5 {
	color: #FF0000;
	font-size: 18px;
	font-family: "黑体";
}
.STYLE6 {
	font-size: 18;
	color: #FF0000;
}
-->
</style>





<div class="listLeft" style="height:auto;">
  <div class="div3" style="height:auto;">
		  <table width="98%" border="0" cellspacing="0" cellpadding="0" style="margin:auto; text-align:left; margin-top:10px; margin-bottom:20px;">
            <form name="order1" id="order1" action="order.asp?Action=Left" method="post" onsubmit="return CheckOrder1(); ">
            <tr>
              <td width="120" height="30">产品名称：</td>
              <td height="30" colspan="3">倍洛加 <span class="STYLE5">（货到付款的在此栏订购）</span></td>
            </tr>
            <tr>
              <td height="30">订购时间：</td>
              <td height="30" colspan="3"><%=FormatDate(Now(),4)%><input type="hidden" name="dgtime" value="<%=Now()%>">
              <%
			Dim THISO:THISO=str&right(year(now),1)&month(now)&day(now)&XXL(5)
			OrderId = HaveOrderId(str,THISO)
			%>
              <input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>"></td> 
            </tr>
            <tr>
              <td height="30">收 货 人：</td>
              <td height="30" colspan="3"><input name="Sh_Name" type="text" class="input4" size="20" id="Sh_Name" />
              （请填写真实姓名）＊＊ </td>
            </tr>
            <!--<tr>
              <td height="30">手 &nbsp;&nbsp;&nbsp;&nbsp; 机：</td>
              <td width="33%" height="30"><input name="Sh_Mobel" type="text" class="input4" size="15" id="Sh_Mobel" /></td>
              <td width="6%" height="30">&nbsp;</td>
              <td width="46%" height="30">&nbsp;</td>
            </tr>-->
            <tr>
              <td height="30">联系电话：</td>
              <td height="30" colspan="3"><input name="Sh_Tel" type="text" class="input4" size="15" id="Sh_Tel" /></td>
            </tr>
            <tr>
              <td height="30">&nbsp;</td>
              <td height="30" colspan="3" style="line-height:20px;"><span class="STYLE6">（注意！请留随身携带的移动电话号码，如：手机、大（小）灵通，这样才便于快递公司送货时及时与您取得联系）</span></td>
            </tr>
            <tr>
              <td height="30">订购数量：</td>
              <td height="30" colspan="3">&nbsp;</td>
            </tr>
            <tr>
              <td height="30" colspan="4" style="line-height:25px; padding-left:20px;">
              	<%Call ProdList()%>              </td>
            <tr>
              <td height="30" colspan="4" align="center"><span class="STYLE1">注：免费快递送货上门</span></td>
            </tr>
            <tr>
              <td height="30">收货地址：</td>
              <td height="30" colspan="3"><input name="Sheng" type="text" class="input4" id="Sheng" size="8" />
              省(如果是直辖市可不填)
              <input name="shi" type="text" class="input4" id="shi" size="8" />
              市<br />
              <input name="xian" type="text" class="input4" id="xian" size="16" />
              <input type="radio" value="1" name="QuType" id="QuType" checked="checked"/>区<input type="radio" name="QuType" id="QuType" value="0" />
              县<span style="margin-left:0px;">（请正确选择区或县）<br />
              <input name="Addres" type="text" class="input4" id="Addres" size="26" onkeyup="this.value=this.value.replace('(','（');this.value=this.value.replace(')','）')" />
              （请填写真实联系地址）＊＊
              </span></td>
            </tr>
            <tr>
              <td height="30">邮政编码：</td>
              <td height="30" colspan="3"><input type="text" name="ZipCode" id="ZipCode" size="6" maxlength="6" class="input4"  />              </td>
            </tr>
            <tr>
              <td height="30">送货方式：</td>
              <td height="30" colspan="3">免费快递送货上门</td>
            </tr>
            <tr>
              <td height="30">支付方式：</td>
              <td height="30" colspan="3">货到付款 (<span class="STYLE3"><font color="#FF0000">银行汇款有优惠</font></span>)</td>
            </tr>
            <tr>
              <td height="80" colspan="4" align="center" valign="bottom"><input id="cod" name="tijiao" type="image" src="images/btn05.jpg"  /></td>
            </tr>
            </form>
          </table>
  </div>
	  </div>
      
<%
    sub ProdList()
    	dim rs,sql
        set rs=server.CreateObject("adodb.recordset")
		sql="select ProductName,Price,Price2,PriceText from NwebCn_Products where ViewFlag=1 order by px asc"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("暂无产品信息，不能定购！")
			response.Write("<input type='hidden' name='RecordCount' id='RecordCount' value='0'>")
		else
			dim i
			i=1
    		while not rs.eof 
				response.Write(rs("ProductName")&"：")
				response.Write("<label>")
					response.Write("<select name='Numbers"&i&"' size='1' id='Numbers"&i&"'>")
						
						response.Write("<option value='NULL' selected>请选择数量</option>")
						response.Write("<option value='"&rs("ProductName")&"(0)'>零盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(1)'>一盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(2)'>二盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(3)'>三盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(4)'>四盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(5)'>五盒</option>")
					response.Write("</select>")
				response.Write("</label>")
				response.Write("【<font color='#ff0000'>"&rs("Price")&rs("PriceText")&"</font><font color='#ff0000'>免费快递送货</font>】")
				response.Write("<br />")
				rs.movenext
				i=i+1
			wend
			response.Write("<input type='hidden' name='RecordCount' id='RecordCount' value='"&rs.recordcount&"'>")
		end if
		rs.close()
		set rs=Nothing
    End sub
    %>
	<script language="javascript">
	<!--
		function CheckOrder1()
		{
			var dgtime,Sh_Name,Sh_Mobel,Sh_Tel,Sheng,shi,xian,Addres,RecordCount,ZipCode
			Sh_Name=document.getElementById("Sh_Name");
			//Sh_Mobel=document.getElementById("Sh_Mobel");
			ShTel=document.getElementById("Sh_Tel");
			//Sheng=document.getElementById("Sheng");
			shi=document.getElementById("shi");
			xian=document.getElementById("xian");
			Addres=document.getElementById("Addres");
			ZipCode=document.getElementById("ZipCode");
			RecordCount=document.getElementById("RecordCount");
			
		if(Sh_Name.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写收货人信息！");
			Sh_Name.focus();
			return false;
		}
		
		
		/*if(Sheng.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写省份信息！");
			Sheng.focus();
			return false;
		}*/
		
		if(shi.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写市级信息！");
			shi.focus();
			return false;
		}
		if(xian.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写县级信息！");
			xian.focus();
			return false;
		}
		if(Addres.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写收货人地址信息！");
			Addres.focus();
			return false;
		}
		if(ZipCode.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写邮编！");
			ZipCode.focus();
			return false;
		}
		
		if(parseInt(RecordCount.value)<=0)
		{
			alert("暂无商品信息，无法继续！")
			return false;
		}
		else
		{
			var falg=false;
			for(var i=1;i<=parseInt(RecordCount.value);i++)
			{
				if(document.getElementById("Numbers"+i).value!="NULL")
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