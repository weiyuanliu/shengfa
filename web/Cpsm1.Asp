<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　　七日科技企业网站管理系统（LiSuo）　　　　　　　  ┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'   版权所有　qisehu.com
'   程序制作　七日科技网络有限公司
'　 Add:四川省成都市二环路西三段181号13楼20/21号
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%>
<% Option Explicit %>
<% response.charset="gb2312" %>
<!--#include file="../Include/NoSqlHack.asp" -->
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
dim rs,sql,SiteTitle,SiteUrl,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,IcpNumber,MesViewFlag,syimg,gonggao,ybpz,qq,syjs
set rs = server.createobject("adodb.recordset")
sql="select top 1 * from NwebCn_Site"
rs.open sql,conn,1,1
SiteTitle=rs("SiteTitle")
SiteUrl=rs("SiteUrl")
ComName=rs("ComName")
Address=rs("Address")
ZipCode=rs("ZipCode")
Telephone=rs("Telephone")
Fax=rs("Fax")
Email=rs("Email")
Keywords=rs("Keywords")
Descriptions=rs("Descriptions")
IcpNumber=rs("IcpNumber")
MesViewFlag=rs("MesViewFlag")
syimg=rs("syimg")
gonggao=rs("Gonggao")
ybpz=rs("ybpz")
QQ=RS("QQ")
syjs=rs("syjs")
rs.close
set rs=nothing '
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=siteTitle%></title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<META NAME="Keywords" CONTENT="<% =Keywords %>" />
<META NAME="Description" CONTENT="<% =Descriptions %>" />
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #010101;
}
body,td,th {
	font-size: 12px;
	color: #ababab;
}
.STYLE1 {
	font-size: 14px;
	font-weight: bold;
	color: #CDD206;
}
.STYLE2 {color: #FF0000}
.STYLE4 {color: #FF0000; font-weight: bold; }
-->
</style></head>

<body>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="Images/Top.jpg" width="1000" height="111" /></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/dhbg.jpg">
  <tr>
    <td><img src="Images/dhleft.jpg" width="63" height="55" /></td>
    <td width="150" align="center"><a href="Index.Asp"><img src="Images/wzsy.jpg" width="81" height="55" border="0" /></a></td>
    <td width="17" align="center"><img src="Images/dhfg.jpg" width="17" height="55" /></td>
    <td width="150" align="center"><a href="Cpsm.Asp" target="_blank"><img src="Images/dhcpsm.jpg" width="91" height="55" border="0" /></a></td>
    <td width="17" align="center"><img src="Images/dhfg.jpg" width="17" height="55" /></td>
    <td width="144" align="center"><a href="LiuYan.Asp" target="_blank"><img src="Images/dhkfly.jpg" width="85" height="55" border="0" /></a></td>
    <td width="17" align="center"><img src="Images/dhfg.jpg" width="17" height="55" /></td>
    <td width="144" align="center"><a href="Wtjd.Asp" target="_blank"><img src="Images/dhwtjd.jpg" width="88" height="55" border="0" /></a></td>
    <td width="17" align="center"><img src="Images/dhfg.jpg" width="17" height="55" /></td>
    <td width="150" align="center"><a href="../zxdg.asp" ><img src="Images/dhzxdg.jpg" width="90" height="55" border="0" /></a></td>
    <td width="62" align="right"><img src="Images/dhright.jpg" width="45" height="55" /></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#000000" style="padding-top:2px; padding-bottom:2px"><img src="Images/banner.jpg" width="1000" height="213" /></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#010101">
  <tr>
    <td width="300" valign="top"><!--#include file="left.asp"--></td>
    <td width="700" align="right" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td style="line-height:22px"><img src="Images/nrongzj.jpg" width="687" height="58" /></td>
      </tr>
      <tr>
        <td style="line-height:22px"><img src="Images/t1.jpg" width="687" height="147" /></td>
      </tr>
      <tr>
        <td style="line-height:22px">


            　　倍洛加虽然受到许多早泄患者的追捧，但是这里给大家申明：倍洛加本身并不是为早泄患者开发的治疗药. 
            而是为性功能正常(只要性交时间超过两分钟均不属于早泄)，对自己性交时间不满意的, 
            想追求更高性生活品质的男性朋友开发的性娱乐产品，倍洛加的目的是让大家都能在床上最彻底最彻底地征服属于你的女人。 
              <br>
              <span class="STYLE4">倍洛加产品说明:              </span><br>
              <span class="STYLE2">【品名】</span> 倍洛加 
            <br>
            <span class="STYLE2">【功效】</span>          倍洛加是纯粹的延时产品。倍洛加任何型号延时效果均为10到45分钟。保底延时时间10分钟。 
          （注：延长时间为使用倍洛加后和平时相比增加时间）注：倍洛加是纯粹的延时产品，所以对因为生理上原因完全不能勃起的患者没有效果。对因为心理原因不能勃起的患者（对自己性能力没信心，对早泄的恐惧等等原因造成的不能勃起）情况，有心理辅助作用。<br> 
          <span class="STYLE2">【规格】</span>          一代产品每盒一支5ml, 一次使用约为0.1ml左右, 一支根据各人个体差别大概可使用50次左右。 倍洛加二代每盒一支10ml, 
          一次使用约为0.1ml左右, 一支根据各人个体差别大概可使用50--100次左右。 
          <br>
          <span class="STYLE2">【使用说明】</span> 外用，按压，外用，按压，一次约0.1毫升左右，将液体喷于手掌中心，然后再用手掌搓揉龟头和冠状沟部位，把液体主要涂抹在龟头和冠状沟部位，阴jing杆上也可以适当涂抹，用手轻揉均匀，促进吸收，15--30分钟后奏效（效力可持续2小时左右）。<br>
          <span class="STYLE2">【注意事项】</span>置阴凉干燥处保存。皮肤破损不宜使用。避免接触眼睛和其他黏膜（如口、鼻等）。本品性状发生改变时禁止使用。请将本品放在儿童不能接触的地方。
          <br>
          <span class="STYLE2">【不良反映】</span>倍洛加在世界范围内使用九年，几乎没有报道过有任何不良反映。但不排除有极个别情况，所以使用过程中如果有过敏症状，如涂抹部位如有烧灼感、红肿等情况应停用，并将局部用清水洗净。 
          <br>
          <br>
          <span class="STYLE2">下面给大家介绍倍洛加的升级产品：倍洛加二代          </span> <br>
          倍洛加二代在延时效果方面和老倍洛加一代产品是没有区别的，升级的地方在于倍洛加二代的配方里加入了刺激女性阴道内敏感区（G点）的成分，当有效成分通过阴jing传递到阴道内的敏感区（G点）的嫩肉后，能很好的刺激其中丰富的交感神经节，特别是能深入某些常规手段很难企及的隐秘区域，从而导致女人的性神经敏感指数持续叠加，女性的阴道肌肉逐渐收缩，随着阴jing在阴道里抽插的持续，绝大多数女性的子宫会有节律性收缩，并从子宫底部一直发展到子宫颈，此时，只要阴jing在阴道里的抽插稍稍加剧，那女方随之而来便是一种极度快感 
          ，这种快感会由阴di开始，并向整个下腹部放射（阴 
          部），随后，一种温暖的浪潮便会从阴部流向全身。最后，女方会感到阴部肌肉发生痉挛和抽搐，有少数的女性在这种极度快感中，还往往会从阴道深处情不自禁地向外喷射类似于男性精液的液体。 
          也就是说，倍洛加二代具多重的功效，除了对男性一贯的优秀延时效果外，又能让性伴侣极易产生性兴奋，达到性高潮。（这个产品就不是大众推广产品了，只推荐给有一定经济实力，对性生活有更高追求的朋友购买使用） 
          .</td>
      </tr>
      <tr>
        <td align="center" style="line-height:22px"><table width="90%" border="0" cellpadding="5" cellspacing="1" bgcolor="#010101">
            <tr>
              <td width="50%" align="center"><img src="images/yidai.jpg"></td>
              <td width="50%" align="center"><img src="images/2dai.jpg"></td>
              </tr>
            <tr>
              <td width="50%" align="center">倍洛加一代</td>
              <td width="50%" align="center">倍洛加二代</td>
              </tr>
            <tr>
              <td width="50%" align="center">110元/盒（1支）</td>
              <td width="50%" align="center">300元/盒（1支）</td>
              </tr>
            <tr>
              <td width="50%" align="center"><a href="zxdg.Asp" style="color:#FFFF00; text-decoration:none">点击订购</a></td>
              <td width="50%" align="center"><a href="zxdg.Asp" style="color:#FFFF00; text-decoration:none">点击订购</a></td>
              </tr>
          </table>
          <br>
          <table width="90%" border="0" cellpadding="5" cellspacing="1" bgcolor="#010101">
            <tr>
              <td width="50%" align="center"><img src="images/yidaijin.jpg" width="200" height="249"></td>
              <td width="50%" align="center"><img src="images/2daijing.jpg" width="200" height="249"></td>
            </tr>
            <tr>
              <td width="50%" align="center">倍洛加一代精包装</td>
              <td width="50%" align="center">倍洛加二代精包装</td>
            </tr>
            <tr>
              <td width="50%" align="center">160元/盒（1支）</td>
              <td width="50%" align="center">350元/盒（1支）</td>
            </tr>
            <tr>
              <td width="50%" align="center"><a href="zxdg.Asp" style="color:#FFFF00; text-decoration:none">点击订购</a></td>
              <td width="50%" align="center"><a href="zxdg.Asp" style="color:#FFFF00; text-decoration:none">点击订购</a></td>
            </tr>
          </table>
          <p>&nbsp;</p></td>
      </tr>
  
      
    </table>
      <span style="line-height:22px"><img src="Images/092.jpg" width="687" height="121" /></span></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/downbg.jpg">
  <tr>
    <td width="34" rowspan="2"><img src="Images/downleft.jpg" width="34" height="98" /></td>
    <td width="933" height="26" align="center">版权所有 &copy; 倍洛加中国官方网站 地址:<%=Address%>ICP备案：粤ICP备06121108号</td>
    <td width="33" rowspan="2"><img src="Images/downright.jpg" width="33" height="98" /></td>
  </tr>
  <tr>
    <td align="center"><table width="80%%" border="0" cellpadding="4" cellspacing="0">
        <tr>
          <td><img src="Images/bottom_aaa.gif" width="120" height="36"></td>
          <td><img src="Images/bottom_other3.gif" width="120" height="36"></td>
          <td><img src="Images/bottom_wenmingchengxin.gif" width="120" height="36"></td>
          <td><img src="Images/cool.gif" width="120" height="36"></td>
          <td><table width="90%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td width="42%"><img src="Images/ind36.gif" width="35" height="43"></td>
                <td width="58%" style="line-height:18px">经营性网站<br>
                备案信息</td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<div style="display:none"><script src="http://s11.cnzz.com/stat.php?id=1109194&web_id=1109194&show=pic" language="JavaScript" charset="gb2312"></script>
</div>
</body>
</html>
