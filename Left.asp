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
            <td height="25" align="right" valign="top" class="lh_24"><a href="News.asp?SortId=58&SortPath=0,58," class="a3">���� &gt;&gt; </a>&nbsp;</td>
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
            <td height="215" valign="top" class="lh_22 cl_013974"><span class="fw_bd">��֤����ӵ���α�����Ч�İ취��</span><br />
              ����ӵ����гɷ־��������޺����ڳ�����֪��١����ȣ����������һ���������ֱ�����Ȼ��������ͷ���֮������Ӵ���Һ�塣ʮ���Ժ󣬽Ӵ����������о��������񾭷�ӳ�������˷�Ӧ��Сʱ���ҿ�ʼ���𽥻ָ�������ע����ͷ���10��󼴿��µ�����Ҫ����Һ���븹�У������û�з�ӳ��Ϊ�ٻ���</td>
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
            <td height="117" valign="top" class="lh_22 cl_013974"><strong>�����žͿɹ����������������˵������۹⣡</strong><br />
              &nbsp;&nbsp;&nbsp;&nbsp;רҵ��׼��װ,�ʵ�Ա��������֪�������ʲô,���˿���Ҳ��Ϊ����Ʒ�С�����Ϊ������˽��ȫ�����˹���,û����֪��������Ǳ���ӡ�<br />
              &nbsp;&nbsp;&nbsp;&nbsp;���ﲻ������ȫ��������,���ڼ�������˽ⱶ���,�������,�����ڲ�����ɫ�������������۷硣</td>
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
              &nbsp;&nbsp;&nbsp;&nbsp;  �����ǿ���Ʒ������ʹ��һЩ�Ƿ���������������,�������ø���;��ɢ���˴�����������ӵ�ҥ��,��Щ��ʮ�ֱ��ɿɳܵģ�<br />
&nbsp;&nbsp;&nbsp;&nbsp; �������ȫ�����۶���,ÿ��ݼ�3000����,����֤�������ǿ��Ĺ�Ч�Լ������ȶ��İ�ȫ��,�����ɷ���ʹ�á�</td>
          </tr>
          <tr>
            <td height="4"></td>
          </tr>
        </table></td>
      </tr>
    </table></td>