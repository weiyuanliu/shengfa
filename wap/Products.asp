<!--#Include file="Head.asp"-->
<!--#include file="../info/ProdClass.asp"-->
<%
Dim Prod
Set Prod=New ProdClass
Prod.Set_TextID(10)
Prod.Set_TableName("NwebCn_About")
%>
	</div>
  <div id="main">
    <div class="topad1"><img src="images/io_tops.jpg" /></div>
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