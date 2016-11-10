<%
Dim Msg
Set Msg=New MsgClass
Msg.Set_HomeNumbers(20)

%>
<div class="boxRight">
	     <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="10"><img src="Images/a3.jpg" width="10" height="43" /></td>
                <td align="left" class="text4" style="background:url(Images/a4.jpg)">
                <span style="float:left; font-size:15px;">【<strong>客户留言</strong>】</span>
                <span style="float:right; margin-right:10px; background-image:url(Images/f12.jpg); width:122px; height:24px; color:#FFFFFF; padding-left:5px;">>>> <a href="Msg.asp"><font color="#FFFFFF">查看更多留言</font></a></span>
                </td>
                <td width="10"><img src="Images/a5.jpg" width="10" height="43" /></td>
              </tr>
              <tr><td height="10"></td></tr>
            </table>
		<div class="box1" style="border:#0351A4 1px solid;">
          	<div style="height:500px;overflow:auto; text-align:center;">
              <div style="width:93%; text-align:left;">
                <ul class="list">
                  <%=Msg.TuiJianList%>
                </ul>
              </div>
              
       	  </div>
		  
		</div>
</div>

<script language="javascript">
<!--
function  Check_MessageValue()
{
	var Msg_Title,Msg_Content,Linkman;
	Msg_Title=document.getElementById("Msg_Title");
	Msg_Content=document.getElementById("Msg_Content");
	Linkman=document.getElementById("Linkman");
	
	if(Msg_Title.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写留言标题！");
		Msg_Title.focus();
		return false;
	}
	if(Msg_Content.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写留言内容！");
		Msg_Content.focus();
		return false;
	}
	if(Linkman.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写留言地区！");
		Linkman.focus();
		return false;
	}
	return true;	
}

function Rest()
{
	var Msg_Title,Msg_Content,Linkman;
	Msg_Title=document.getElementById("Msg_Title");
	Msg_Content=document.getElementById("Msg_Content");
	Linkman=document.getElementById("Linkman");
	Msg_Title.value="";
	Msg_Content.value="";
	Linkman.value="";
	return false;
}
-->
</script>
