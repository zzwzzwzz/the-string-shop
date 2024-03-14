<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_info_OnOff,root_info_OffNote,root_info_LogoPic,root_info_skin from root_info where id=1"
rs.open sql,conn,1,1
root_info_OnOff           =rs(0)
root_info_OffNote         =rs(1)
root_info_LogoPic         =rs(2)
root_info_skin            =rs(3)
rs.close
set rs=nothing

if root_info_skin="" then
    response.write "<link href=style/default.css rel=stylesheet type=text/css>"
else
    response.write "<link href=style/"&root_info_skin&".css rel=stylesheet type=text/css>"
end if

if root_info_OnOff=1 then 
    response.write "<center><br><br><br><br><br>"&root_info_OffNote&"</center>"
    response.end
end if
%>
<script src="js/PicLimit.js" type="text/javascript"></script>
<!--header begin-->
<%
response.write  "<div id=mainbox>"&_
				"<table border=0 width=100% cellpadding=4 style='border-collapse: collapse' class='top_table'>"&_
				"<form name=form_search action=Product_ListSearch.asp method=get>"&_
				"	<tr><td colspan=2 height=5></td></tr>"&_
				"	<tr>"&_
				"		<td><img src=uploadpic/"&root_info_LogoPic&"></td>"&_
				"		<td align=right>"&_
				"			<table><tr><td class='cartimg'>&nbsp;&nbsp;&nbsp;</td><td><a href=Cart_List.asp>购物车</a>"
							if session("y")<>"" then response.write "(<font color=#FF0000>"&session("y")&"</font>)"
response.write  "			| <a href=User_Fav.asp>收藏夹</a> | <a href=User_reg.asp>注册</a> | <a href=User_login.asp>登录</a> | <a href=admin/Index.asp target=_blank>后台管理</a>"&_
				"			</td></tr></table><br>"&_
 				"			商品搜索: <input type=text size=30 name=name> <select name=bid size=1>"&_
				"			<option value=''>所有类别</option>"
							sql="select prod_bigclass_id,prod_bigclass_name from prod_bigclass order by prod_bigclass_id desc"
							set rs=conn.execute (sql)
							set prod_bigclass_id=rs(0)
							set prod_bigclass_name=rs(1)
							do while not rs.eof
response.write  "			<option value="&prod_bigclass_id&">"&prod_bigclass_name&"</option>"
							rs.movenext
							loop
							rs.close
							set rs=nothing
response.write  "			</select> <input class=button type=submit value=搜索>&nbsp;&nbsp; <a href=Product_Search.asp>高级搜索</a>"&_
				"		</td>"&_
				"	</tr>"&_
				"	<tr><td colspan=2 height=5></td></tr>"&_
				"	<tr>"&_
				"		<td colspan=2 class=TopMenu>"&_
				"			<a href=Index.asp class=M>网站首页</a>&nbsp;&nbsp; |&nbsp;&nbsp;"&_ 
				"			<a href=Product_ListFlag.asp?flag=1 class=M>全部商品</a>&nbsp;&nbsp; |&nbsp;&nbsp;"&_ 
				"			<a href=News_List.asp class=M>文章中心</a>&nbsp;&nbsp; |&nbsp;&nbsp; "&_
				"			<a href=User_Index.asp class=M>个人中心</a>&nbsp;&nbsp; |&nbsp;&nbsp; "&_
				"			<a href=GuestBook_List.asp class=M>留言评论</a>&nbsp;&nbsp; |&nbsp;&nbsp; "&_
				"			<a href=Help_List.asp class=M>帮助中心</a>&nbsp;&nbsp; |&nbsp;&nbsp; "&_
				"			<a href=ContactUs.asp class=M>联系我们</a></td>"&_
				"	</tr>"&_
				"</form>"&_
				"</table>"&_
				"<div class=brclass></div>"
%><!--header end-->
