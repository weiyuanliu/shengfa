<!--#Include file="Head.asp"-->
<!--#include file="info/MsgClass.asp"-->
<!--#include file="../Include/360_safe3.asp" -->
<!--#include file="../ip.asp" -->
<%
	Dim Msg
	Set Msg=New MsgClass
	Msg.Set_Page_Size(15)
	dim sai1,sai2,saihe
	Randomize 
	sai1=int(10*Rnd)
	Randomize 
	sai2=int(10*Rnd)
	saihe=sai1+sai2
	response.cookies("sai1")=sai1
	response.cookies("sai2")=sai2
	response.cookies("vcode")=saihe

%>
	</div>
<div id="main">
 <div class="topad1"><img src="images/io_tops.jpg" /></div>
 <div class="html">
 <div class="feedbacklist">
          <div>
             <div style="border-bottom:1px #CCC solid;padding-bottom:10px;">
             <table width="100%" border="0">
               <tr>
                 <td><%=GetValues("NwebCn_About","Content",62)%></td>
               </tr>
             </table>
             </div>
              <div style="margin-bottom:10px;padding-top:5px;border-bottom:1px #CCC solid;"> 
    			<a name="add"></a>
                <div style="height:0px; line-height:0px;"></div>  
                <form name="SaveMsg" id="SaveMsg" action="SaveMsg.asp" method="post" onSubmit="return Verify(this)">
                <input type="hidden" name="Msg_Addres" id="Msg_Addres" size="30" value="<%=Request.ServerVariables("REMOTE_ADDR")%>"/>
                     <div style="text-align:center;font-size:16px;color:#0417f1;font-weight:bold;line-height:50px">我 要 留 言 </div>
                        <table cellspacing="0" width="100%" border="0" class="addfeedbacktable">
                          <tr>
                            <td width="90" align="center" ><strong>标 &nbsp;  题：</strong></td>
                            <td><input type="text" name="Msg_Title" id="Msg_Title" size="60" maxlength="50" class="feedw" style="width:60%;" />
                              <span >&nbsp;*</span></td>
                          </tr>
                            <tr>
                            <td width="90" align="center" ><strong>地 &nbsp;  区：</strong></td>
                            <td><input name="hip" id="hip" type="text" style="width:60%;" value="自动识别" size="20" readonly >
                            </td>
                          </tr>
                            <tr>
                            <td width="90" align="center" ><strong>验证码：</strong></td>
                            <td><input name="yzm" id="yzm" type="text" size="10" style="width:20%;" />
                                <span><img src="../Include/newcode1226.asp" alt="验证码看不清楚?请点击刷新验证码!" title="验证码看不清楚?请点击刷新验证码!" style="cursor:pointer;width:auto;margin-bottom:-6px;" onClick="this.src='../Include/newcode1226.asp?t='+(new Date().getTime());" />&nbsp;*验证码区分大小写</span></td>
                          </tr>
                          <tr>
                            <td height="90" align="center"><strong>内 &nbsp;  容：</strong></td>
                            <td><textarea name="Msg_Content" rows="5" cols="60" id="Msg_Content" class="feedw" style="width:80%;" /></textarea>
                                <span>&nbsp;*</span></td>
                          </tr>
                          <tr>
                            <td height="42" colspan="2" align="center"><input name="" type="image" src="images/feedback1.jpg"  alt="提交" class="feedbacksub"/>
                               <input name="" type="image" src="images/feedback2.jpg" alt="复位" onClick="this.form.reset();return false;" class="feedbacksub"/>
                              </td>
                           </tr>
                          </tbody>
                        </table>
                        <input type="hidden" name="action" value="add" />
                    </form>
         </div>
        </div>
             <div class="feedbacktable">
             <ul> 
                     <%=Msg.List%>
             </ul>
             </div>
    </div>          
   </div>
   <div></div>
 </div>
<!-- Verify Script Start -->
<script language="JavaScript">
  function checkMail(str){var pattern = /^([a-zA-Z0-9_-_.])+@([a-zA-Z0-9_-])+(\.[a-zA-Z0-9_-])+/;if(pattern.test(str)) return true; else return false}
  function trim(str){var pattern = /(^\s+)$/;str = str.replace(pattern, "");var pattern = /(\s+)$/;str = str.replace(pattern, "");return str;}
  function Verify(frm) {
  	  if (trim(frm.Msg_Title.value) == "")
  	   {
  	   	alert("请输入主题!");
  	   	frm.Msg_Title.focus();
  	   	return false;
  	   }
  	  if (trim(frm.Msg_Content.value) == "")
  	   {
  	   	alert("请输入内容!");
  	   	frm.Msg_Content.focus();
  	   	return false;
  	   }
  	  if (trim(frm.yzm.value) == "")
  	   {
  	   	alert("请输入验证码!");
  	   	frm.yzm.focus();
  	   	return false;
  	   }
  }
</script>
<!-- Verify Script End -->
<!--#Include file="Foot.asp"-->