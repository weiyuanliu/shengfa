<td width="221" valign="top"><table width="212" border="0" cellpadding="0" cellspacing="0" class="bk_zt">
      <tr>
        <td height="73" align="center" class="bk_xb2"><img src="images/phone_sonpage.jpg" /></td>
      </tr>
      <%
	  dim wxtk
	  wxtk=false
	  if wxtk then
	  %>
      <tr>
        <td height="50" align="center"><img src="images/img_29.gif" /></td>
      </tr>
      <%
	  end if
	  %>
    </table>
    <table width="212" border="0" cellpadding="0" cellspacing="0" class="bk_zt mg_t10">
      <tr>
        <td height="44" align="center" class="bk_xb2"><img src="images/img_30.gif" /></td>
      </tr>
      <tr>
        <td height="50" align="center"><img src="images/img_31.gif" /></td>
      </tr>
    </table>
    <table width="212" border="0" cellpadding="0" cellspacing="0" class="bk_zt mg_t10">
      <tr>
        <td height="100" align="center" class="bk_xb2"><%=Guanggao(11,179,89)%></td>
      </tr>
      <tr>
        <td height="100" align="center"><img src="images/img_33.gif" /></td>
      </tr>
    </table>
    <table width="212" border="0" cellspacing="0" cellpadding="0" class="mg_t10">
      <tr>
        <td><img src="images/pic_8.gif" /></td>
      </tr>
      <tr>
        <td class="bk_xb1 bk_zb bk_yb"><table width="200" border="0" cellpadding="0" cellspacing="0" class="mag">
          <tr>
            <td height="215" class="lh_24" valign="top">
            <%
			call NewsList("0,58,")
			Function NewsList(Spath)
			 Dim rs,sql
			 set rs=server.CreateObject("Adodb.recordset")
			 sql="Select top 10 * from NwebCn_News where ViewFlag=1 and Charindex(SortPath,'"&Spath&"')>0 Order by AddTime desc,px desc,id desc"
			 rs.open sql,conn,1,1
			 if not rs.eof then
			  while not rs.eof
			  %>
			    <a href="NewsView.asp?Id=<%=rs("Id")%>&SortId=<%=rs("SortId")%>" title="<%=rs("NewsName")%>" class="a3">&middot;<%=StrLeft(rs("NewsName"),26)%></a><br />
			  <%
			  rs.movenext
			  wend
			 end if
			 rs.close
			 set rs=nothing
			End Function
			%>
          </tr>
          <tr>
            <td height="25" align="right" valign="top" class="lh_24"><a href="News.asp?SortId=58&SortPath=0,58," class="a3">更多 &gt;&gt; </a>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table>
    <table width="212" border="0" cellpadding="0" cellspacing="0" class="bk_zt5 mg_t10">
      <tr>
        <td height="4"></td>
      </tr>
      <tr>
        <td height="100" align="center"><img src="images/img_14.jpg" /></td>
      </tr>
      <tr>
        <td height="100" align="center"><img src="images/img_15.gif" /></td>
      </tr>
      <tr>
        <td height="100" align="center"><img src="images/img_16.gif" /></td>
      </tr>
      <tr>
        <td height="4"></td>
      </tr>
    </table>
    <table width="212" border="0" cellspacing="0" cellpadding="0" class="mg_t10">
      <tr>
        <td><img src="images/pic_34.gif" /></td>
      </tr>
      <tr>
        <td class="bk_xb1 bk_zb bk_yb"><table width="200" border="0" cellpadding="0" cellspacing="0" class="mag">
          <tr>
            <td height="215" valign="top" class="lh_22 cl_013974"><span class="fw_bd">验证倍洛加得真伪最简单有效的办法：</span><br />
              倍洛加得所有成分均对身体无害，口尝即可知真假。首先，将倍洛加喷一点于您的手背处，然后再用舌头五分之二面积接触溶液体。十秒以后，接触面会有特殊感觉，表面神经反映减慢。此反应半小时左右开始会逐渐恢复正常（注：舌头舔后10秒后即可吐掉，不要把溶液吞入腹中），如果没有反映则为假货。</td>
          </tr>
          <tr>
            <td height="4"></td>
          </tr>
        </table></td>
      </tr>
    </table>
    <table width="212" border="0" cellspacing="0" cellpadding="0" class="mg_t10">
      <tr>
        <td><img src="images/pic_35.gif" /></td>
      </tr>
      <tr>
        <td class="bk_xb1 bk_zb bk_yb"><table width="200" border="0" cellpadding="0" cellspacing="0" class="mag">
          <tr>
            <td height="30" valign="top"><img src="images/pic_19.gif" /></td>
          </tr>
          <tr>
            <td height="117" valign="top" class="lh_22 cl_013974"><strong>不出门就可购买，无须再忍受旁人的异样眼光！</strong><br />
              &nbsp;&nbsp;&nbsp;&nbsp;专业标准包装,邮递员都不可能知道里边是什么,别人看到也以为是礼品盒。我们为您的隐私安全做足了功夫,没有人知道你买的是倍洛加。<br />
              &nbsp;&nbsp;&nbsp;&nbsp;购物不出门完全消除尴尬,坐在家里就能了解倍洛加,购买倍洛加,让你在不动声色中重振起男人雄风。</td>
          </tr>
          <tr>
            <td height="4"></td>
          </tr>
        </table></td>
      </tr>
    </table>
    <table width="212" border="0" cellspacing="0" cellpadding="0" class="mg_t10">
      <tr>
        <td><img src="images/pic_36.gif" /></td>
      </tr>
      <tr>
        <td class="bk_xb1 bk_zb bk_yb"><table width="200" border="0" cellpadding="0" cellspacing="0" class="mag">
          <tr>
            <td height="53" align="center" valign="middle"><img src="images/img_37.gif" /></td>
          </tr>
          <tr>
            <td height="117" valign="top" class="lh_22 cl_013974">
              &nbsp;&nbsp;&nbsp;&nbsp;  倍洛加强大的品牌优势使得一些非法竞争者难以生存,他们利用各种途径散布了大量攻击倍洛加的谣言,这些是十分卑鄙可耻的！<br />
&nbsp;&nbsp;&nbsp;&nbsp; 倍洛加在全球销售多年,每年惠及3000万人,足以证明倍洛加强大的功效以及长期稳定的安全性,您尽可放心使用。</td>
          </tr>
          <tr>
            <td height="4"></td>
          </tr>
        </table></td>
      </tr>
    </table></td>