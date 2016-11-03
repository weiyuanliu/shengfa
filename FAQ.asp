<!--#Include file="Head.asp"-->
<!--#include file="info/lxwmClass.asp"-->
<%
dim About
Set About=New LxwmClass
About.Set_TableName("NwebCn_About")
About.Set_ID(49)
%>
     <div style="background:url(style/blue/images/datu.gif) center  no-repeat; width:1420px;height:410px;margin:0 auto;"></div>
     <div style="background:url(style/blue/images/header_05.jpg) center  no-repeat; width:1420px;height:111px;margin:0 auto;"></div>

	</div>
  <div id="main">
      <div class="topad1"><img src="style/blue/images/news_top.jpg" width="988" height="66" /></div>
      <div class="html">
        <div class="listfaq">
          <%=About.PrintText%>
        </div>
     </div>
     <div></div>
  </div>
<!--#Include file="Foot.asp"-->