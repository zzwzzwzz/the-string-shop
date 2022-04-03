<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
select case action
	case "wadd"
   		call wadd()
 	case "modifyz"
   		call modifyz()
 	case "del"
   		call del()
 	case ""
   		case else
   	response.write "操作错误"
end select

sub wadd()
	prod_brand=my_request("prod_brand",0)
  	if prod_brand="" then
    	response.write"<SCRIPT language=JavaScript>alert('商品品牌名称不能为空');"
    	response.write"javascript:history.go(-1)</SCRIPT>"
    	response.end
  	end if
  	sql="insert into prod_brand (prod_brand) values('"&prod_brand&"')"
  	conn.execute(sql)
  	response.redirect "product_brand_list.asp"
end sub


sub modifyz()
  	id=my_request("id",1)
  	prod_brand=my_request("prod_brand",0)
  	if prod_brand="" or id="" then
    	response.write"<SCRIPT language=JavaScript>alert('信息未填写完整或操作错误');"
    	response.write"javascript:history.go(-1)</SCRIPT>"
    	response.end
  	end if
  	sql="update prod_brand set prod_brand='"&prod_brand&"' where id="&id
  	conn.execute(sql)
  	response.redirect "product_brand_list.asp"
end sub

sub del()
  	id=my_request("id",1)
  	sql = "delete from prod_brand where id="&id
  	conn.execute(sql)
  	response.redirect "product_brand_list.asp"
end sub

%>

<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品品牌-管理</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="3" class="header">商品品牌-管理</td>
	</tr>
	<tr>
		<td><b>品牌名称</b></td>
		<td colspan=2><b>操作区</b></td>
	</tr>
	<%
		sql="select * from prod_brand"
		set rs=server.CreateObject ("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.eof then
			response.write "<tr><td colspan=3>暂时还没有任何数据，请添加</td></tr>"
		else
			i=1
			do while not rs.eof
			set id=rs("id")
			set prod_brand=rs("prod_brand")
		%>
		<form action="product_brand_list.asp">
			<input type="hidden" name="action" value="modifyz">
			<input type="hidden" name="id" value="<%=id%>">
		<tr>
			<td>
			<input type="text" name="prod_brand" size="17" maxlength="50" value="<%=prod_brand%>"></td>
			<td><input type="submit" value="修改保存" name="B4"></td>
			<td><a href="product_brand_list.asp?id=<%=id%>&action=del">删除</a></td>
		</tr>
		</form>
		<%
			rs.movenext
			i=i+1
			loop
		end if
		rs.close
		set rs=nothing
		%>
</tbody>
</table>
<br>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
		<form action="product_brand_list.asp?action=wadd" method=post>
		<tr>
			<td class="header">添加新品牌名称</td>
		</tr>
		<tr>
			<td colspan="4">添加新品牌名称：<input type="text" name="prod_brand" size="20" maxlength="50"> 
			<input type="submit" value="添加" name="B1">
			<input type="reset" value="重置" name="B3"></td>
		</tr>
		</form>
</tbody>
</table>

</body>

</html>