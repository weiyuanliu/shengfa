<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><hr size="1" color="#9DBBDB" /></td>
  </tr>
  <tr>
    <td height="50" class="text3">��Ȩ���� &copy; <%=ComName%>  <%=IcpNumber%>
 ��ַ:<%=Address%> &nbsp;<a href="TestLinks.asp" target="_blank">��������</a>
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
*    ������Ϣ����  
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
*    ��Ϣж���¼���������д  
*/  
CLASS_MSN_MESSAGE.prototype.onunload = function() {  
    return true;  
}  
/* 
*    ��Ϣ�����¼���Ҫʵ���Լ������ӣ�����д��  
*  
*/  
CLASS_MSN_MESSAGE.prototype.oncommand = function(){  
    //this.close = true;
    this.hide();  
 //window.open("http://www.baidu.com");
   
} 
/**//*  
*    ��Ϣ��ʾ����  
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
        str += "<SPAN title=�ر� style='FONT-WEIGHT: bold; FONT-SIZE: 12px; CURSOR: hand; COLOR: red; MARGIN-RIGHT: 4px' id='btSysClose' >��</SPAN></TD>"  
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
** �����ٶȷ��� 
**/ 
CLASS_MSN_MESSAGE.prototype.speed = function(s){ 
    var t = 20; 
    try { 
        t = praseInt(s); 
    } catch(e){} 
    this.speed = t; 
} 
/**//* 
** ���ò������� 
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
var MSG1 = new CLASS_MSN_MESSAGE("aa",380,480,"��Ϣ��ʾ��","<table border=0 cellpadding=0 cellspacing=0 width='100%'><tr><td align='center' style='color:#FF0000; font-size:12px;'>����Ǵ������̰���<br/>������վʵʩ��թ��Υ����Ϊ</td></tr></table>","<table border=0 cellpadding=0 cellspacing=0 width='100%'><tr><td align='center' style='font-size:12px; text-align:left; line-height:16px;'>&nbsp;&nbsp;&nbsp;&nbsp;���ڱ����ǿ��Ĺ�Ч�Լ����۵Ļ𱬴�����ͬ�е����棻������������һЩ�������ӵ�ע�⡣�����������ǽ����ӵ���թ�绰�����������в��������������ṩ�Ŀ��ϻ���Ҫ�������ϴ��������������Ǳ���Ӳ�Ʒ��˵Ҫ�����Ǳ���Ӳ�Ʒ����ɨ�ء�������Ȼ���������ӽ������������ϴ���������һЩ���⹥�����̰�����Ӳ�Ʒ�����ӡ����ɶ����һЩ��������һ������թ����������Զ�����ν��˽��ð��315��վ��Ҳ���⽫��Щ�̰���ת����������վ���棬Ȼ���绰������Ҫ����������ǧ��������Щ��ν�ġ�˽��ð��315��վ�������������ǽ��̰���ɾ��������ð��315��վ�ڰٶ�һ�����Ͷ���ţë��������֮�����ǹ���������վ�����һЩ��ҵ�������Ͷ������Ȼ���绰�������ҵ����֧��һ��Ǯ�ͽ���̰���ɾ�����������繥���̰�ȷʵ����һ���ĸ���Ӱ�죬��֮����άȨĿǰȷʵ�Ƚ����ѣ�����һЩ��ҵ�����Ʋ����֣���͸������ġ�˽��ð��315��վ���㹻������ռ䣬ʹ���𽥷��Ĳ�Ᵽ���<br>&nbsp;&nbsp;&nbsp;&nbsp;����վ�ڴ��ٴξ��治�����ӣ�����Ӳ�Ʒ�ڹ�����������������ʵʵ���ڵ�Ч��Ӯ���˹�������ߵ����Ρ����ͻ��ڿ��ഫʹ���Ǳ����ӵ���˴�������ʵ�û���һЩͬ���Լ��������ӱ����һЩ�����ֶβ����������Ӵ���ʲô����Ӱ�죬һ������Ĳ�Ʒ���ǲ������ӷ������̰������ܱ�Ĩ�ڵģ����⣬Ŀǰ�����Ѿ������������Ѿ��������������������ӷ�����IP��ַ�����Ų��ú󲻷����Ӿͻ�Ϊ��Ƿ�������ҥ�̰�����Ϊ�е���Ӧ���������Ρ�</td></tr><tr><td align='right' style='font-size:12px;'>����������ӹ���������޹�˾<br>2009��11��16��</br></td></tr></table>");  
    MSG1.rect(null,null,null,screen.height-50); 
    MSG1.speed    = 10; 
    MSG1.step    = 5; 
    MSG1.show();  
	-->
</SCRIPT> 
<%end if%>