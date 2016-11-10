<!--#Include file="Head.asp"-->
<%
Dim Action:Action=request("Action")
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" valign="top">
    
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="100%" height="50" valign="top" class="titlebg"><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="6" height="6"></td>
              <td width="141" ></td>
              <td rowspan="2" align="right" class="cl_fff" style="padding-right:10px;"><a href="/wap/" class="a1">首页</a> > <a href="AliPay.asp" class="a1">支付宝购买</a></td>
            </tr>
            <tr>
              <td height="38"></td>
              <td align="center" class="fz_24 fw_bd cl_013974 title">支付宝购买</td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td valign="top" class="bk_xb1 bk_zb bk_yb pd6">
          <div style="clear:both; padding:6px; border:#0250A2 1px solid;background:#D0EAF7; text-align:left; line-height:25px; margin-bottom:10px;">
<div style="text-align:center; color:#FF0000; font-size:16px; padding-top:5px; padding-bottom:5px;"><strong><%=GetValues("NwebCn_About","AboutName",59)%></strong></div>
<span style="font-size:14px;"><%=GetValues("NwebCn_About","Content",59)%></span>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" name="dinggou" id="dinggou">
  <tr>
    <td valign="top" align="center">
    
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="10"><img src="../Images/a3.jpg" width="10" height="43" /></td>
        <td align="left" class="text4" style="background:url(../Images/a4.jpg)"><strong>支付宝付款订购（先付款有优惠）</strong></td>
        <td width="10"><img src="../Images/a5.jpg" width="10" height="43" /></td>
      </tr>
    </table>
    
    
    </td>
  </tr>
  <tr>
  	<td colspan="2">
    <table border="0" cellpadding="0" cellspacing="0" align="center">
    <tr>
    <td width="100%">
      <%if Action ="ArrayPlay" then%>		
      	<!--#include file="info/OnArrayPlay.asp"-->
      <%else%>
      	<!--#include file="info/ArrayPlay.asp"-->
      <%end if%>
      </td>
      </tr>
      </table>
    </td>
  </tr>
</table>

          </td>
        </tr>
    </table></td>
  </tr>
</table>
<!--#Include file="Foot.asp"-->