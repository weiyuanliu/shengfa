<!--#Include file="Head.asp"-->
<!--#include file="info/ProdClass.asp"-->
<%
Dim Prod
Set Prod=New ProdClass
Prod.Set_TextID(10)
Prod.Set_TableName("NwebCn_About")
%>
     <div style="background:url(style/blue/images/datu.gif) center  no-repeat; width:1420px;height:410px;margin:0 auto;"></div>
     <div style="background:url(style/blue/images/header_05.jpg) center  no-repeat; width:1420px;height:111px;margin:0 auto;"></div>

	</div>
  <div id="main">
    <div class="topad1"><img src="style/blue/images/news_top.jpg" width="988" height="66" /></div>
    <div class="html">
        <div class="html1 htmlcss">
        <%=Prod.Text%>
        </div>
            <div id="ProdList" style="text-align:center;">
            	<%
		Prod.SetPicture_Width(300)
		Prod.SetPicture_Height(300)
		Prod.ProdList
		%>
      		</div>
     </div>
   </div>
<!--#Include file="Foot.asp"-->