<!--#Include file="Head.asp"-->
<!--#include file="../info/lxwmClass.asp"-->
<%
dim About
Set About=New LxwmClass
About.Set_TableName("NwebCn_About")
About.Set_ID(63)
%>
       <div class="topad1"><img src="images/io_tops.jpg" /></div>
	</div>
    <div id="main">
      <div class="html">
        <div class="html1 htmlcss">
<%=About.PrintText%>
</div>
      </div>
   </div>
<!--#Include file="Foot.asp"-->