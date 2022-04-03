<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select top 1 root_info_QQOnOff from root_info"
rs.open sql,conn,1,1
root_info_QQOnOff =rs(0)
rs.close
set rs=nothing

if root_info_QQOnOff=0 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select top 1 root_info_QQ,root_info_QQplace,root_info_QQName,root_info_WangWang,root_info_WangWangOnOff from root_info"
	rs.open sql,conn,1,1
	root_info_QQ =rs(0)
	root_info_QQplace=rs(1)
	root_info_QQName=rs(2)
	root_info_WangWang=rs(3)
	root_info_WangWangOnOff=rs(4)
	rs.close
	set rs=nothing
	
	QQ=Split(root_info_QQ,",")
	QQName=Split(root_info_QQName,",")
	if root_info_QQplace<>2 and root_info_QQplace<>3 then
%>
<script type="text/javascript">
//<![CDATA[
var tips; var theTop = 110/*这是默认高度,越大越往下*/; 
var old = theTop;

window.onload=function initFloatTips() {
  tips = document.getElementById('floatTips');
  moveTips();
};

function moveTips() {
  var tt=40;
  if (window.innerHeight) {
    pos = window.pageYOffset
  }
  else if (document.documentElement && document.documentElement.scrollTop) {
    pos = document.documentElement.scrollTop
  }
  else if (document.body) {
    pos = document.body.scrollTop;
  }
  pos=pos-tips.offsetTop+theTop;
  pos=tips.offsetTop+pos/10;
  if (pos < theTop) pos = theTop;
  if (pos != old) {
    tips.style.top = pos+"px";
    tt=10;
  }
  old = pos;
  setTimeout(moveTips,tt);
}

function showDiv() { 
var oDiv=document.getElementById("floatTips"); 
oDiv.style.display=(oDiv.style.display=="none")?"block":"none"; 
} 

function closeDL(){
	ifCouplet=false;
	document.getElementById('floatTips').style.visibility='hidden';
}
//!]]>
</script>
<%if root_info_QQPlace=0 then%>
<style type="text/css">
div#floatTips{
 position:absolute;
 border:solid 0px #777;
 padding:3px;
 left:0px;
 top:250px;
 width:100px;
 color:white;
}
</style>
<%end if%>
<%if root_info_QQPlace=1 then%>
<style type="text/css">
div#floatTips{
 position:absolute;
 padding:3px;
 right:8px;
 top:250px;
 width:100px;
 color:white;
}
</style>
<%end if%>
<div id="floatTips">
<map name="FPMap0"><area shape="rect" coords="6, 8, 50, 21" href="#" onclick="javascript:closeDL();return false;" target="_self"></map>
<table border="0" cellspacing="0" cellpadding="0">
<tr><td><img border=0 src=images/kefu_up.gif usemap="#FPMap0"></td></tr>
<tr><td><table width=100% border="0" background=images/kefu_middle.gif cellspacing="0" cellpadding="0">
<%
for N=0 to UBound(QQ)
%>
<tr>
  <td valign=middle background=images/kefu_middle.gif height="23">&nbsp;&nbsp;
  	<a target=_blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=trim(qq(n))%>&Site=在线咨询&Menu=no title=<%=trim(qqname(n))%>><img src=http://wpa.qq.com/pa?p=1:<%=trim(qq(n))%>:4 border=0 align=absmiddle>&nbsp;<%=trim(qqname(n))%></a>
  </td>
</tr>
<%next%>
<%if root_info_WangWangOnOff=0 then%>
<tr><td height="26">&nbsp;
	<SCRIPT language="JavaScript" src="js/URLEncode.js"></SCRIPT>
	<script language="javascript">
	var taobaoid;
	var taobaos;
    taobaos="<%=root_info_WangWang%>";
    taobaoid=URLEncode(taobaos)
    document.writeln("<a target=_blank href=http://amos1.taobao.com/msg.ww?v=2&s=1&uid="+taobaoid+">")
	document.writeln("<img border=0 alt=点击这里给我发消息 src=http://amos1.taobao.com/online.ww?v=2&s=1&uid="+taobaoid+">")
	document.writeln("</a>")
	</script></td></tr>
  <%end if%>
</table></td></tr>
<tr><td><img border=0 src=images/kefu_down.jpg></td></tr>
</table>

</div>
<%end if
end if%>