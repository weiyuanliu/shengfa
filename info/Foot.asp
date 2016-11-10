<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><hr size="1" color="#9DBBDB" /></td>
  </tr>
  <tr>
    <td height="50" class="text3">版权所有 &copy; <%=ComName%>  <%=IcpNumber%>
 地址:<%=Address%> &nbsp;<a href="TestLinks.asp" target="_blank">友情链接</a>
</td>
  </tr>
</table>
<%

Dim ViewMsg
ViewMsg=false
if ViewMsg then
%>
<SCRIPT language=JavaScript>  
<!--  
function CLASS_MSN_MESSAGE(id,width,height,caption,title,message,target,action){  
    this.id     = id;  
    this.title  = title;  
    this.caption= caption;  
    this.message= message;  
    this.target = target;  
    this.action = action;  
    this.width    = width?width:200;  
    this.height = height?height:120;  
    this.timeout= 1000;  
    this.speed    = 20; 
    this.step    = 1; 
    this.right    = screen.width -1;  
    this.bottom = screen.height; 
    this.left    = this.right - this.width; 
    this.top    = this.bottom - this.height; 
    this.timer    = 0; 
    this.pause    = false;
    this.close    = false;
    this.autoHide    = true;
}  
  
/*
*    隐藏消息方法  
*/  
CLASS_MSN_MESSAGE.prototype.hide = function(){  
    if(this.onunload()){  
        var offset  = this.height>this.bottom-this.top?this.height:this.bottom-this.top; 
        var me  = this;  
        if(this.timer>0){   
            window.clearInterval(me.timer);  
        }  
        var fun = function(){  
            if(me.pause==false||me.close){
                var x  = me.left; 
                var y  = 0; 
                var width = me.width; 
                var height = 0; 
                if(me.offset>0){ 
                    height = me.offset; 
                } 
     
                y  = me.bottom - height; 
     
                if(y>=me.bottom){ 
                    window.clearInterval(me.timer);  
                    me.Pop.hide();  
                } else { 
                    me.offset = me.offset - me.step;  
                } 
                me.Pop.show(x,y,width,height);    
            }             
        }  
        this.timer = window.setInterval(fun,this.speed)      
    }  
}  
  
/*  
*    消息卸载事件，可以重写  
*/  
CLASS_MSN_MESSAGE.prototype.onunload = function() {  
    return true;  
}  
/* 
*    消息命令事件，要实现自己的连接，请重写它  
*  
*/  
CLASS_MSN_MESSAGE.prototype.oncommand = function(){  
    //this.close = true;
    this.hide();  
 //window.open("http://www.baidu.com");
   
} 
/**//*  
*    消息显示方法  
*/  
CLASS_MSN_MESSAGE.prototype.show = function(){  
    var oPopup = window.createPopup(); //IE5.5+  
    this.Pop = oPopup;  
  
    var w = this.width;  
    var h = this.height;  
  
    var str = "<DIV style='BORDER-RIGHT: #455690 1px solid; BORDER-TOP: #a6b4cf 1px solid; Z-INDEX: 99999; LEFT: 0px; BORDER-LEFT: #a6b4cf 1px solid; WIDTH: " + w + "px; BORDER-BOTTOM: #455690 1px solid; POSITION: absolute; TOP: 0px; HEIGHT: " + h + "px; BACKGROUND-COLOR: #c9d3f3'>"  
        str += "<TABLE style='BORDER-TOP: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid' cellSpacing=0 cellPadding=0 width='100%' bgColor=#cfdef4 border=0>"  
        str += "<TR>"  
        str += "<TD style='FONT-SIZE: 12px;COLOR: #0f2c8c' width=30 height=24></TD>"  
        str += "<TD style='PADDING-LEFT: 4px; FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #1f336b; PADDING-TOP: 4px' vAlign=center width='100%'>" + this.caption + "</TD>"  
        str += "<TD style='PADDING-RIGHT: 2px; PADDING-TOP: 2px' vAlign=center align=right width=19>"  
        str += "<SPAN title=关闭 style='FONT-WEIGHT: bold; FONT-SIZE: 12px; CURSOR: hand; COLOR: red; MARGIN-RIGHT: 4px' id='btSysClose' >×</SPAN></TD>"  
        str += "</TR>"  
        str += "<TR>"  
        str += "<TD style='PADDING-RIGHT: 1px;PADDING-BOTTOM: 1px' colSpan=3>"  
        str += "<DIV style='BORDER-RIGHT: #b9c9ef 1px solid; PADDING-RIGHT: 8px; BORDER-TOP: #728eb8 1px solid; PADDING-LEFT: 8px; FONT-SIZE: 12px; PADDING-BOTTOM: 8px; BORDER-LEFT: #728eb8 1px solid; WIDTH: 100%; COLOR: #1f336b; PADDING-TOP: 8px; BORDER-BOTTOM: #b9c9ef 1px solid; HEIGHT: 100%'>" + this.title + "<BR><BR>"  
        str += "<DIV style='WORD-BREAK: break-all; margin-top:-20px;' align=left><A href='javascript:void(0)' hidefocus=false id='btCommand'><FONT color=#ff0000>" + this.message + "</FONT></A><A href='http:' hidefocus=false id='ommand'></A></DIV>"  
        str += "</DIV>"  
        str += "</TD>"  
        str += "</TR>"  
        str += "</TABLE>"  
        str += "</DIV>"  
  
    oPopup.document.body.innerHTML = str;   
    this.offset  = 0; 
    var me  = this;  
    oPopup.document.body.onmouseover = function(){me.pause=true;}
    oPopup.document.body.onmouseout = function(){me.pause=false;}
    var fun = function(){  
        var x  = me.left; 
        var y  = 0; 
        var width    = me.width; 
        var height    = me.height; 
            if(me.offset>me.height){ 
                height = me.height; 
            } else { 
                height = me.offset; 
            } 
        y  = me.bottom - me.offset; 
        if(y<=me.top){ 
            me.timeout--; 
            if(me.timeout==0){ 
                window.clearInterval(me.timer);  
                if(me.autoHide){
                    me.hide(); 
                }
            } 
        } else { 
            me.offset = me.offset + me.step; 
        } 
        me.Pop.show(x,y,width,height);    
    }  
  
    this.timer = window.setInterval(fun,this.speed)      
    var btClose = oPopup.document.getElementById("btSysClose");  
    btClose.onclick = function(){  
        me.close = true;
        me.hide();  
    }  
  
    var btCommand = oPopup.document.getElementById("btCommand");  
    btCommand.onclick = function(){  
        me.oncommand();  
    }    
  var ommand = oPopup.document.getElementById("ommand");  
      ommand.onclick = function(){  
       //this.close = true;
    me.hide();  
 window.open(ommand.href);
    }   
}  
/**//* 
** 设置速度方法 
**/ 
CLASS_MSN_MESSAGE.prototype.speed = function(s){ 
    var t = 20; 
    try { 
        t = praseInt(s); 
    } catch(e){} 
    this.speed = t; 
} 
/**//* 
** 设置步长方法 
**/ 
CLASS_MSN_MESSAGE.prototype.step = function(s){ 
    var t = 1; 
    try { 
        t = praseInt(s); 
    } catch(e){} 
    this.step = t; 
} 
  
CLASS_MSN_MESSAGE.prototype.rect = function(left,right,top,bottom){ 
    try { 
        this.left        = left    !=null?left:this.right-this.width; 
        this.right        = right    !=null?right:this.left +this.width; 
        this.bottom        = bottom!=null?(bottom>screen.height?screen.height:bottom):screen.height; 
        this.top        = top    !=null?top:this.bottom - this.height; 
    } catch(e){} 
} 
var MSG1 = new CLASS_MSN_MESSAGE("aa",380,480,"消息提示：","<table border=0 cellpadding=0 cellspacing=0 width='100%'><tr><td align='center' style='color:#FF0000; font-size:12px;'>严厉谴责发网络诽谤贴<br/>对我网站实施敲诈的违法行为</td></tr></table>","<table border=0 cellpadding=0 cellspacing=0 width='100%'><tr><td align='center' style='font-size:12px; text-align:left; line-height:16px;'>&nbsp;&nbsp;&nbsp;&nbsp;由于倍洛加强大的功效以及销售的火爆触及到同行的利益；甚至于引起了一些不法分子的注意。近半月来我们接连接到敲诈电话，犯罪分子威胁我们如果不向其提供的卡上汇款就要在网络上大量发贴攻击我们倍洛加产品，说要让我们倍洛加产品声名扫地。果不其然，不法分子近几日在网络上大量发布了一些恶意攻击、诽谤倍洛加产品的帖子。更可恶的是一些网络上面一向以敲诈勒索而臭名远扬的所谓“私人冒牌315网站”也故意将这些诽谤贴转贴在他们网站上面，然后打电话给我们要付款三至五千给他们这些所谓的“私人冒牌315网站”，才能替我们将诽谤贴删除（这种冒牌315网站在百度一搜索就多如牛毛，其生财之道就是故意在其网站上面对一些企业发起恶意投诉贴，然后打电话给相关企业让其支付一笔钱就将其诽谤贴删除，由于网络攻击诽谤确实具有一定的负面影响，加之网络维权目前确实比较困难，所以一些企业往往破财消灾，这就给大量的“私人冒牌315网站”足够的生存空间，使其逐渐泛滥猖獗）。<br>&nbsp;&nbsp;&nbsp;&nbsp;我网站在此再次警告不法分子，倍洛加产品在国内销售两年来，以实实在在的效果赢得了广大消费者的信任。广大客户口口相传使我们倍洛加拥有了大量的忠实用户，一些同行以及不法分子背后的一些卑鄙手段并不会给倍洛加带来什么负面影响，一个优秀的产品岂是不法分子发几个诽谤贴就能被抹黑的，另外，目前我们已经报案，网警已经立案正在锁定不法分子发贴的IP地址，相信不久后不法分子就会为其非法攻击造谣诽谤的行为承担相应的刑事责任。</td></tr><tr><td align='right' style='font-size:12px;'>西班牙倍洛加国际香港有限公司<br>2009年11月16日</br></td></tr></table>");  
    MSG1.rect(null,null,null,screen.height-50); 
    MSG1.speed    = 10; 
    MSG1.step    = 5; 
    MSG1.show();  
	-->
</SCRIPT> 
<%end if%>