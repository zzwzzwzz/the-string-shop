<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品高级搜索</title>
<link rel="stylesheet" type="text/css" href="style.css">
<%
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from prod_smallclass order by prod_smallclass_bid desc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
subcat[0] = new Array("此大类下所有小类","<%= trim(rs("prod_smallclass_bid"))%>","");
        <%
        count = 1
        do while not rs.eof 
        ss=trim(rs("prod_smallclass_bid"))
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("prod_smallclass_name"))%>","<%= trim(rs("prod_smallclass_bid"))%>","<%= trim(rs("prod_smallclass_id"))%>");
        <%
        count = count + 1
        rs.movenext
        if trim(rs("prod_smallclass_bid"))<>ss then
        %>
subcat[<%=count%>] = new Array("此大类下所有小类","<%= trim(rs("prod_smallclass_bid"))%>","");   
        <%
        count = count + 1
        end if
        loop
        rs.close
        set rs=nothing
        %>
onecount=<%=count%>;

//类别切换
function changelocation(locationid)
    {
    document.form1.sid.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form1.sid.options[document.form1.sid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    } 
</script>
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Product_Info_List.asp" method="get" name="form1">
    <tr>
		<td colspan="2" class="title">商品高级搜索</td>
	</tr>
	<tr>
		<td>商品名称：</td>
		<td><input type="text" name="KeyWord" size="20"></td>
	</tr>
	<tr>
		<td>所属类别：</td>
		<td>
			<select name="bid" onChange="changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)">
		    <option value="">请选择大类</option>
		    <%
		    sql="select * from prod_bigclass order by prod_bigclass_id desc"
		    set rs=conn.execute (sql)
		    do while not rs.eof
		    %>
		    <option value="<%=rs("prod_bigclass_id")%>"><%=rs("prod_bigclass_name")%></option>
		    <%
		    rs.movenext
		    loop
		    rs.close
		    set rs=nothing
		    %>
            </select><select name="sid"> 
            <option value="">请选择小类</option>
            </select>
        </td>
	</tr>
	<tr>
		<td>商品内容含：</td>
		<td><input type="text" name="product_info_Detail" size="20"></td>
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
			<input type="reset" value="  重  置  " name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>

