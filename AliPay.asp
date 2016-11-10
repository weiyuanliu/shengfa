<!--#Include file="Head.asp"-->
<%
Dim Action:Action=request("Action")
%>
     <div style="background:url(style/blue/images/header_03.jpg) center  no-repeat; width:1420px;height:334px;margin:0 auto;"></div>
     <div style="background:url(style/blue/images/header_05.jpg) center  no-repeat; width:1420px;height:111px;margin:0 auto;"></div>
	</div>
  <div id="main">
    <div class="topad1"><img src="style/blue/images/order_tops.jpg" width="988" height="66" /></div>
<SCRIPT src="style/blue/js/alipay_submit.js" type="text/javascript"></SCRIPT>
    <div class="html">
    <div class="html1">
     <div class="listrightb">
      <%if Action ="ArrayPlay" then%>		
      	<!--#include file="info/OnArrayPlay.asp"-->
      <%else%>
      	<!--#include file="info/ArrayPlay.asp"-->
      <%end if%>
       </div>
     </div>
   <div>
  </div>
  </div>
  </div>
<!--#Include file="Foot.asp"-->