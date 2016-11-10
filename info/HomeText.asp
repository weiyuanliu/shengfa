<%
Dim Text
Set Text=New TextClass
Text.Set_ID(51)
Text.Set_ZiDuan("Content")
Text.Set_TableName("NwebCn_About")
%>
<div class="boxLeft">
    <%=Text.Print_Text%>
</div>