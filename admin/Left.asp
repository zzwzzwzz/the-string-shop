<!--#include file="admin_check.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>左侧导航</title>
<link rel="stylesheet" type="text/css" href="style.css">
<base target="main">
</head>

<body>
<script language="JavaScript">
function ClearAllDeploy(){
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var userdeploy='';
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		userdeploy=deployitem.substring(0,admin_start);
	}
	for(i=0;i<20;i++){
		obj=document.getElementById("cate_"+"id"+i);	
		img=document.getElementById("img_"+"id"+i);
		if(obj && obj.style.display=="none"){
			obj.style.display="";
			img_re=new RegExp("_open\\.gif$");
			img.src=img.src.replace(img_re,'_fold.gif');
		}
	}
	deployitem=userdeploy+"\n\t\t";
	SetCookie("deploy",deployitem);
}
function SetAllDeploy(){
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var userdeploy='';
	var admindeploy='';
	var i;
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		userdeploy=deployitem.substring(0,admin_start);
	}
	for(i=0;i<20;i++){
		obj=document.getElementById("cate_"+"id"+i);	
		img=document.getElementById("img_"+"id"+i);
		if(obj && obj.style.display==""){
			obj.style.display="none";
			img_re=new RegExp("_fold\\.gif$");
			img.src=img.src.replace(img_re,'_open.gif');
		}
		admindeploy=admindeploy+"id"+i+"\t";
	}
	deployitem=userdeploy+"\n\t"+admindeploy;
	SetCookie("deploy",deployitem);
}
function IndexDeploy(ID,type){
	obj=document.getElementById("cate_"+ID);	
	img=document.getElementById("img_"+ID);
	if(obj.style.display=="none"){
		obj.style.display="";
		img_re=new RegExp("_open\\.gif$");
		img.src=img.src.replace(img_re,'_fold.gif');
		SaveDeploy(ID,type,false);
	}else{
		obj.style.display="none";
		img_re=new RegExp("_fold\\.gif$");
		img.src=img.src.replace(img_re,'_open.gif');
		SaveDeploy(ID,type,true);
	}
	return false;
}
function SaveDeploy(ID,type,is){
	var foo=new Array();
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var admindeploy='';
	var userdeploy='';
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		admindeploy= deployitem.substring(admin_start+1,deployitem.length);
		userdeploy = deployitem.substring(0,admin_start);
	}
	if(deployitem!=null){
		if(admin_start!=-1){
			deployitem = type==0 ? userdeploy : admindeploy;
		}
		deployitem=deployitem.split("\t");
		for(i in deployitem){
			if(deployitem[i]!=ID && deployitem[i]!=""){
				foo[foo.length]=deployitem[i];
			}
		}
	}
	if(is){
		foo[foo.length]=ID;
	}
	deployitem = type==0 ? "\t"+foo.join("\t")+"\t\n"+admindeploy : userdeploy+"\n\t"+foo.join("\t")+"\t";
	SetCookie("deploy",deployitem)
}
function SetCookie(name,value){
	expires=new Date();
	expires.setTime(expires.getTime()+(86400*365));
	document.cookie=name+"="+escape(value)+"; expires="+expires.toGMTString()+"; path=/";
}
function FetchCookie(name){
	var start=document.cookie.indexOf(name);
	var end=document.cookie.indexOf(";",start);
	return start==-1 ? null : unescape(document.cookie.substring(start+name.length+1,(end>start ? end : document.cookie.length)));
}
</script>


<table border="0" width="100%" cellpadding="4" style="border:1px solid #CCCCCC; border-collapse: collapse; padding-left:4px; padding-right:4px; padding-top:1px; padding-bottom:1px" bgcolor="#FFFFFF">
	<tr>
		<td align="center"><a onClick="return ClearAllDeploy()" href="#">+ 展开菜单</a>&nbsp;&nbsp; &nbsp;<a onClick="return SetAllDeploy()" href="#">- 关闭</a></td>
	</tr>
	<tr>
		<td align="center"><a href="Right.asp" target="main">后台首页</a> <font color="#999999">| </font>&nbsp;<a href="Admin_LoginOut.asp" target=_parent>退出管理</a></td>
	</tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td height="1"></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id1',1)"><img id="img_id1" src="images/cate_open.gif" border=0></a>
		  <a href="#" onclick="return IndexDeploy('id1',1)" class="a_black">基本设置</td>
	</tr>
	<tbody id="cate_id1" style="display:none;">
	<tr>
		<td class="altbg2">
		   <li><a target="main" href="Root_Info_Set.asp">基本资料设置</a></li>
		   <li><a target="main" href="Root_Model_list.asp">网站模板设置</a></li>
		   <li><a target="main" href="Root_AboutUs_Set.asp">关于我们设置</a></li>
		   <li><a target="main" href="Root_Option_Set.asp">参数选项设置</a></li>
		   <li><a href="Root_NetPay_Set.asp" target="main">在线支付设置</a></li>
		   <li><a href="Root_remit_set.asp" target="main">汇款说明设置</a></li>
		   <li><a href="Root_Deliver_Set.asp" target="main">送货方式设置</a><br></li>
		   <li><a href="Root_Email_set.asp" target="main">群发邮箱设置</a></li>
		   <li><a href="Root_Vote_set.asp" target="main">投票调查设置</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id2',1)"><img id="img_id2" src="images/cate_open.gif" border=0></a>
		  <a href="#" onclick="return IndexDeploy('id2',1)" class="a_black">商品管理</td>
	</tr>
	<tbody id="cate_id2" style="display:none;">
	<tr>
		<td class="altbg2">
		   <li><a target="main" href="Prod_Class_List.asp">类别管理</a></li>
		   <li><a target="main" href="Product_Info_List.asp">商品管理</a> | 
			<a target="main" href="Product_Info_Add.asp">添加</a></li>
		   <li><a target="main" href="Product_Info_Search.asp">商品高级搜索</a></li>
		   <li><a target="main" href="Product_Brand_list.asp">商品品牌管理</a></li>
		   <li><a target="main" href="Product_kucun_list.asp">商品库存管理</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id3',1)"><img id="img_id3" src="images/cate_open.gif" border=0></a>
		  <a href="#" onclick="return IndexDeploy('id3',1)" class="a_black">订单管理</td>
	</tr>
	<tbody id="cate_id3" style="display:none;">
	<tr>
		<td class="altbg2">
		   <li><a href="Order_info_List.asp" target="main">订单管理</a></li>
		   <li><a href="Order_info_search.asp" target="main">订单高级搜索</a></li>
		   <li><a href="Order_info_recycle.asp" target="main">订单回收站</a></li>
		   <li><a href="Order_info_Print.asp" target="main">订单打印</a></li>
		   <li><a href="Order_info_SaleCount.asp" target="main">销售统计</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id10',1)"><img id="img_id10" src="images/cate_open.gif" border=0></a>
		  <a href="#" onClick="return IndexDeploy('id10',1)" class="a_black">会员管理</a></td>
	</tr>
	<tbody id="cate_id10" style="display:none;">
	<tr>
		<td class="altbg2">
		   <li><a href="user_option_set.asp" target="main">会员选项设置</a></li>
		   <li><a href="user_level_list.asp" target="main">会员级别管理</a></li>
		   <li><a href="user_info_list.asp" target="main">会员信息管理</a></li>
		   <li><a href="user_info_search.asp" target="main">会员高级搜索</a></li>
		   <li><a href="user_email_list.asp" target="main">会员邮件群发</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id4',1)"><img id="img_id4" src="images/cate_open.gif" border=0></a>
		<a href="#" onClick="return IndexDeploy('id4',1)" class="a_black">新闻管理</a></td>
	</tr>
	<tbody id="cate_id4" style="display:none;">
	<tr>
		<td class="altbg2">
		   <li><a target="main" href="News_Info_Add.asp">新闻动态添加</a></li>
		   <li><a href="news_info_list.asp" target="main">新闻动态管理</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id5',1)"><img id="img_id5" src="images/cate_open.gif" border=0></a>
		  <a href="#" onClick="return IndexDeploy('id5',1)" class="a_black">留言及评论</a></td>
	</tr>
	<tbody id="cate_id5" style="display:none;">
	<tr>
		<td class="altbg2">
		    <li><a target="main" href="GB_Info_List.asp">在线留言管理</a></li>
  			<li><a target="main" href="Prod_Review_List.asp"> 商品评论管理</a></li>		   
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>


<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id8',1)"><img id="img_id8" src="images/cate_open.gif" border=0></a>
		  <a href="#" onClick="return IndexDeploy('id8',1)" class="a_black">帮助中心</a></td>
	</tr>
	<tbody id="cate_id8" style="display:none;">
	<tr>
		<td class="altbg2">
		   <li><a href="help_info_add.asp" target="main">帮助信息添加</a></li>
		   <li><a href="help_info_list.asp" target="main">帮助信息管理</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>


<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id11',1)"><img id="img_id11" src="images/cate_open.gif" border=0></a>
		  <a href="#" onClick="return IndexDeploy('id11',1)" class="a_black">管理权限</a></td>
	</tr>
	<tbody id="cate_id11" style="display:none;">
	<tr>
		<td class="altbg2">
		   <li><a href="admin_info_add.asp" target="main">管理人员添加</a></li>
		   <li><a href="admin_info_list.asp" target="main">管理人员管理</a></li>
		   <li><a href="admin_info_PassWordModiByUserName.asp?admin_info_UserName=<%=session("admin_info_UserName")%>" target="main">管理密码修改</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header"><a style="float:right" href="#" onClick="return IndexDeploy('id9',1)"><img id="img_id9" src="images/cate_open.gif" border=0></a>
		  <a href="#" onClick="return IndexDeploy('id9',1)" class="a_black">友情链接</a></td>
	</tr>
	<tbody id="cate_id9" style="display:none;">
	<tr>
		<td class="altbg2">
		   <li><a href="link_info_add.asp" target="main">友情链接添加</a></li>
		   <li><a href="link_info_list.asp" target="main">友情链接管理</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td>　</td>
	</tr>
</table>


<table border="1" width="100%" cellpadding="4" style="border-collapse: collapse; padding-left:4px; padding-right:4px; padding-top:1px; padding-bottom:1px" bgcolor="#FFFFCC" bordercolor="#808080">
	<tr>
		<td align="center">深度网上购物系统<br>
		开发商：<a target="_blank" href="http://www.deepne.cn/">深度网络</a></td>
	</tr>
</table>
</body>

</html>

