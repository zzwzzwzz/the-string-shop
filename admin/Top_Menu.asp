<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>demo</title>


<style type="text/css">
<!--
body {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
}
#nav, #nav ul {
    float: left;
	list-style: none;
	line-height: 22px;
	background: #F9F9F9;
	font-weight: bold;--
	padding: 0px;
	margin: 0px;
	border: solid 1px #CCCCCC;
	border-right: 0px;

}

#nav ul ul{
	border: solid 1px #CCCCCC;
	border-top: 0px;
	border-right: 0px;
}

#nav a {
	display: block;
	width: 80px;
	color: #333333;
	text-decoration: none;
	text-align: center;
	border-right: solid 1px #CCCCCC;
}

#nav a:hover{
	color: #336666;
}

#nav li {
	float: left;
	width: 80px;
}

#nav li ul {
	position: absolute;
	left: -999em;
	width: 100px;
	font-weight: normal;
	margin: 0px;
	padding: 0px;
}

#nav li li {
	width: 100px;
}

#nav li ul a {
	width: 100px;
	padding: 0px 12px;
	line-height: 19px;
	border-top: solid 1px #CCCCCC;
	text-align: left;
}

#nav li ul ul {
	margin: -20px 0 0 99px; 
}

#nav li:hover ul ul,#nav li.sfhover ul ul{
	left: -999em;
}

#nav li:hover ul, #nav li li:hover ul,#nav li.sfhover ul, #nav li li.sfhover ul{
	left: auto;
}

#nav li:hover, #nav li.sfhover {
	background: #DDE3E9;
}
-->
</style>

<script type="text/javascript"><!--//--><![CDATA[//><!--

sfHover = function() {
	var sfEls = document.getElementById("nav").getElementsByTagName("LI");
	for (var i=0; i<sfEls.length; i++) {
		sfEls[i].onmouseover=function() {
			this.className+=" sfhover";
		}
		sfEls[i].onmouseout=function() {
			this.className=this.className.replace(new RegExp(" sfhover\\b"), "");
		}
	}
}
if (window.attachEvent) window.attachEvent("onload", sfHover);

//--><!]]></script>
</head>
<body>
<!--manu div start-->
<ul id="nav">
    <!-- root -->
	<li><a href="Root_Info_set.asp">基本设置</a>
		<ul>
			<li><a href="Root_Info_set.asp"		>基本资料设置</a></li>
  			<li><a href="Root_Option_set.asp"	>选项参数设置</a></li>
    		<li><a href="Root_Transport_set.asp">配送方式设置</a></li>
    		<li><a href="Root_Pay_set.asp"		>付款方式设置</a></li>
    		<li><a href="Root_Vote_set.asp"		>投票调查设置</a></li>
		</ul>
	</li>
    <!-- product -->
	<li><a href="Product_Info_list.asp">商品管理</a>
		<ul>
			<li><a href="Product_Class_Add.asp" >商品类别添加</a></li>
			<li><a href="Product_Class_list.asp">商品类别管理</a></li>
			<li><a href="Product_Info_Add.asp"  >商品信息添加</a></li>
			<li><a href="Product_Info_List.asp" >商品信息管理</a></li>
			<li><a href="Product_Info_Search.asp" >商品高级搜索</a></li>
		</ul>
	</li>
    <!-- order -->
	<li><a href="Order_Info_List.asp">订单管理</a>
		<ul>
			<li><a href="Order_Info_List.asp"  >订单信息管理</a></li>
			<li><a href="Order_Info_Search.asp">订单信息查询</a></li>
		</ul>
	</li>
    <!-- news -->
	<li><a href="News_Info_List.asp">文章管理</a>
	</li>
	<!-- help -->
	<li><a href="Help_Info_AboutUs.asp">帮助管理</a>
		<ul>
			<li><a href="Help_Info_AboutUs.asp"  >关于我们设置</a></li>
			<li><a href="Help_Info_ContactUs.asp">联系我们设置</a></li>
			<li><a href="Help_Info_Pay.asp"		 >付款说明设置</a></li>
			<li><a href="Help_Info_Transport.asp">运送说明设置</a></li>
			<li><a href="Help_Info_FAQ.asp"		 >常见问题解答</a></li>
		</ul>
	</li>
	<!-- guestbook -->
	<li><a href="GuestBook_Info_List.asp">留言管理</a>
	</li>
	<!-- link -->
	<li><a href="Link_Info_List.asp">友情链接</a>
	</li>
	<!-- manage -->
	<li><a href="Manage_Info_Set.asp">管理设置</a>
	</li>
	<!-- index -->
	<li><a href="Index.asp">管理首页</a>
	</li>
</ul>
<br><br><br>
</body>
</html>

