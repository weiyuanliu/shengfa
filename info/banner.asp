<style type="text/css">
<!--
.STYLE1 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
<div class="banner Clear">
	  <img src="images/banner.jpg" class="imgLeft"  alt="����ӹٷ���վ ��������"  title="����ӹٷ���վ ��������" />
	  <div class="right">
      <form name="search" id="search" method="post" style="margin:0px;" action="OrderSearchlist.asp" onsubmit="return CheckValue2();">
	    <div class="div1 Clear">
		  <input name="KeyWord" type="text" class="input3" value="<%=Trim(Request("KeyWord"))%>"/>
		  <input name="search" type="image" src="images/btn01.jpg" style="position:relative; top:6px; *top:3px" />
		</div>
       </form>
		<div class="div2" style="text-indent:24px;">
        	<span class="STYLE1">�Ѿ��ύ�����󳬹����컹δ�յ������������ڴ˵�������ԣ�            </span>
       	  <div style="text-align:center; margin-top:10px;">
        	<input type="image" src="images/Msgbutton.jpg" onclick="window.open('Msg2.asp');" />
            </div>
        </div>
  </div>
	</div>