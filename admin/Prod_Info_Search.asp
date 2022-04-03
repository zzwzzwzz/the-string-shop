<!--#include file="admin_check.asp"-->
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品高级搜索</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="prod_Info_List.asp" method="get" name="form1">
    <tr>
		<td colspan="2" class="title">商品高级搜索</td>
	</tr>
	<tr>
		<td>商品名称：</td>
		<td><input type="text" name="KeyWord" size="20"></td>
	</tr>
	<tr>
		<td>所属类别：</td>
		<td><select name="cid">
            <option value="">所有类别下</option>
		    <%
		     sql="select cid,prod_class_name from prod_class order by cid desc"
		     set rs=conn.execute (sql)
		     set cid=rs(0)
		     set prod_class_name=rs(1)
		     do while not rs.eof
		    %>
		    <option value="<%=cid%>"><%=prod_class_name%></option>
		    <%
		     rs.movenext
		     loop
		     rs.close
		     set rs=nothing
		    %>
		 </select></td>
	</tr>
	<tr>
		<td>商品内容含：</td>
		<td><input type="text" name="prod_info_Detail" size="20"></td>
	</tr>
	<tr>
		<td>本站价格范围：</td>
		<td><input type="text" name="prod_info_PriceSMin" size="6">元(小值)&nbsp; 至 
		<input type="text" name="prod_info_PriceSMax" size="6">元(大值)</td>
	</tr>
	<tr>
		<td>搜索结果排序：</td>
		<td><input type="radio" CHECKED value="1" name="sort">时间降 
		<input type="radio" value="2" name="sort">时间升 
		<input type="radio" value="3" name="sort">编号降 
		<input type="radio" value="4" name="sort">编号升 
		<input type="radio" value="5" name="sort">商品名称 
		<input type="radio" value="6" name="sort">浏览次数</td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="  搜  索  " name="B1">&nbsp;&nbsp;&nbsp;
			<input type="reset" value="重置" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>

