<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=5
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("prod_review_id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""参数错误!"");location.href=""prod_review_List.asp"";</script>")
  response.end
end if

sql="select prod_review_name,prod_review_pid,prod_review_detail,prod_review_time,prod_review_BackDetail from prod_review where prod_review_id="&id
set rs=conn.execute (sql)
prod_review_name=rs(0)
prod_review_pid=rs(1)
prod_review_detail=rs(2)
prod_review_time=rs(3)
prod_review_BackDetail=rs(4)
rs.close
set rs=nothing

sql="select product_info_name from product_info where id="&prod_review_pid
set rs=conn.execute (sql)
prod_info_name=rs(0)
rs.close
set rs=nothing 

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    id                   =my_request("prod_review_id",1)
    prod_review_detail    =my_request("prod_review_detail",0)
    prod_review_BackDetail=my_request("prod_review_BackDetail",0)

    if id="" or prod_review_detail="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from prod_review where prod_review_id="&id
        rs.open sql,conn,1,3
        rs("prod_review_detail")    =prod_review_detail
        rs("prod_review_BackDetail")=prod_review_BackDetail
        rs("prod_review_BackTime")  =now()
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功回复/更新了一条评论信息！","prod_review_list.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品-评论信息-回复</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="prod_review_back.asp" method="post">
<input type="hidden" name="action" value="save"> 
<input type="hidden" name="prod_review_id" value="<%=id%>"> 
	<tr>
		<td colspan="2" class="header">商品评论信息-回复</td>
	</tr>
	<tr>
		<td>评论人姓名：</td>
		<td><%=prod_review_name%></td>
	</tr>
	<tr>
		<td>评论时间：</td>
		<td><%=prod_review_time%></td>
	</tr>
	<tr>
		<td>评论商品：</td>
		<td><b><a href=../product_detail.asp?id=<%=prod_review_pid%> target="_blank"><%=prod_info_name%></a></b></td>
	</tr>
	<tr>
		<td>评论内容：</td>
		<td><textarea rows="8" name="prod_review_detail" cols="60"><%=prod_review_detail%></textarea></td>
	</tr>
	<tr>
		<td>回复内容：</td>
		<td><textarea rows="8" name="prod_review_BackDetail" cols="60"><%=prod_review_BackDetail%></textarea></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="提交" name="Submit1">&nbsp; 
		   <input type="reset" value="重置" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
 
