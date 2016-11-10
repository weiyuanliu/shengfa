   <div id="footer">
  <div class="corpright">
 公司：古那迪國際集團(香港)有限公司&nbsp;&nbsp;&nbsp;&nbsp;<br>&nbsp;&nbsp;&nbsp;&nbsp;<br>&nbsp; &nbsp; <br>
 </div>
 <div class="sm">郑重声明：未经授权禁止转载、摘编、复制或建立镜像&nbsp;&nbsp;&nbsp;&nbsp;版权 &copy; 古那迪中国官方网站&nbsp; </div>
</div>
   </div>
</div>
</div>

<%if otherscount and gonggao<>"" and (nm="index") then%>
<div id="msg_win" style="display:block;visibility:visible;position:absolute;right:2;opacity:1;">
    <div class="icos">
    <a id="msg_min" title="最小化" href="javascript:void 0">最小化</a> <a>|</a> <a id="msg_close" title="关闭" href="javascript:void 0">关闭提示框</a>
    </div>
    <div id="msg_title"><%=taobaoid%></div>
    <div id="msg_content">
    <%=Or2(gonggao)%>
    </div>
</div>
	<%
	 Dim P
	 P=split(syjs,"|")
	 Dim Pwidth,Pheight
	 Pwidth = P(0)
	 Pheight = P(1)
	%>
    <style type="text/css">
    #msg_win{background:#f3f3f3;width:<%=Pwidth%>px;margin:0px;display:none;overflow:hidden;z-index:999;}
    #msg_win .icos{position:absolute;top:2px;*top:0px;right:2px;z-index:9;}
    .icos a{float:left;color:#833B02;margin:1px;text-align:center;text-decoration:none;font-family:webdings; color:#FFFFFF; font-size:16px;}
    .icos a:hover{color:#fff;}
    #msg_title{background:#013976;border-bottom:1px solid #A67901;color:#FFF;height:25px;line-height:25px;text-indent:5px;}
    #msg_content{margin:2px;width:<%=Pwidth-10%>px;height:<%=Pheight%>px;overflow:hidden; padding:5px; line-height:20px;}
    </style>

    <script language="javascript">
    var Message={
    set: function() {//最小化与恢复状态切换
    var set=this.minbtn.status == 1?[0,1,'block',this.char[0],'最小化']:[1,0,'none',this.char[1],'恢复'];
    this.minbtn.status=set[0];
    this.win.style.borderBottomWidth=set[1];
    this.content.style.display =set[2];
    this.minbtn.innerHTML =set[3]
    this.minbtn.title = set[4];
    this.win.style.top = this.getY().top;
    this.win.style.right = 2;
    },
    close: function() {//关闭
    this.win.style.display = 'none';
    window.onscroll = null;
    },
    setOpacity: function(x) {//设置透明度
    var v = x >= 100 ? '': 'Alpha(opacity=' + x + ')';
    this.win.style.visibility = x<=0?'hidden':'visible';//IE有绝对或相对定位内容不随父透明度变化的bug
    this.win.style.filter = v;
    this.win.style.opacity = x / 100;
    },
    show: function() {//渐显
    clearInterval(this.timer2);
    var me = this,fx = this.fx(0, 100, 0.1),t = 0;
    this.timer2 = setInterval(function() {
    t = fx();
    me.setOpacity(t[0]);
    if (t[1] == 0) {clearInterval(me.timer2) }
    },10);
    },
    fx: function(a, b, c) {//缓冲计算
    var cMath = Math[(a - b) > 0 ? "floor": "ceil"],c = c || 0.1;
    return function() {return [a += cMath((b - a) * c), a - b]}
    },
    getY: function() {//计算移动坐标
    var d = document,b = document.body, e = document.documentElement;
    var s = Math.max(b.scrollTop, e.scrollTop);
    var h = /BackCompat/i.test(document.compatMode)?b.clientHeight:e.clientHeight;
    var h2 = this.win.offsetHeight;
    return {foot: s + h + h2 + 2+'px',top: s + h - h2 - 2+'px'}
    },
    moveTo: function(y) {//移动动画
    clearInterval(this.timer);
    var me = this,a = parseInt(this.win.style.top)||0;
    var fx = this.fx(a, parseInt(y));
    var t = 0 ;
    this.timer = setInterval(function() {
    t = fx();
    me.win.style.top = t[0]+'px';
    me.win.style.right = '2px';
    if (t[1] == 0) {
    clearInterval(me.timer);
    me.bind();
    }
    },10);
    },
    bind:function (){//绑定窗口滚动条与大小变化事件
    var me=this,st,rt;
    window.onscroll = function() {
    clearTimeout(st);
    clearTimeout(me.timer2);
    me.setOpacity(0);
    st = setTimeout(function() {
    me.win.style.top = me.getY().top;
	me.win.style.right = '2px';
    me.show();
    },600);
    };
    window.onresize = function (){
    clearTimeout(rt);
    rt = setTimeout(function() {me.win.style.top = me.getY().top},100);
    }
    },
    init: function() {//创建HTML
    function $(id) {return document.getElementById(id)};
    this.win=$('msg_win');
    var set={minbtn: 'msg_min',closebtn: 'msg_close',title: 'msg_title',content: 'msg_content'};
    for (var Id in set) {this[Id] = $(set[Id])};
    var me = this;
    this.minbtn.onclick = function() {me.set();this.blur()};
    this.closebtn.onclick = function() {me.close()};
    this.char=navigator.userAgent.toLowerCase().indexOf('firefox')+1?['最小化','恢复',' | 关闭提示框']:['0','2','r'];//FF不支持webdings字体
    this.minbtn.innerHTML=this.char[0];
    this.closebtn.innerHTML=this.char[2];
    setTimeout(function() {//初始化最先位置
    me.win.style.display = 'block';
    me.win.style.top = me.getY().foot;
    me.moveTo(me.getY().top);
    },0);
    return this;
    }
    };
    Message.init();
    </script>
<%end if%>


</body>
</html>