<!--#Include file="Head.asp"-->
<!--#include file="../info/lxwmClass.asp"-->
<%
dim About
Set About=New LxwmClass
About.Set_TableName("NwebCn_About")
About.Set_ID(49)
%>
	</div>
  <div id="main">
      <div class="topad1"><img src="images/io_tops.jpg" /></div>
      <div class="html">
        <div class="listfaq">
          <%=About.PrintText%>
        </div>
     </div>
     <div></div>
  </div>
<!--#Include file="Foot.asp"-->