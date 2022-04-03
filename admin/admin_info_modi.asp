<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=9
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("admin_info_id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""参数错误!"");location.href=""help_info_List.asp"";</script>")
  response.end
end if

sql="select admin_info_RealName,admin_info_flag,admin_info_username from admin_info where admin_info_id="&id
set rs=conn.execute (sql)
admin_info_RealName=rs("admin_info_RealName")
admin_info_flag    =rs("admin_info_flag")
admin_info_username=rs("admin_info_username")
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call Save()
end if

//管理人员-修改保存
sub Save() 
    id                  =my_request("admin_info_id",1)
    admin_info_RealName =my_request("admin_info_RealName",0) 
    for i=0 to 9
        b=request(i)
        if b="" then b=0
        a=a&","&b
    next
    a=right(replace(a," ",""),len(replace(a," ",""))-1)

    if admin_info_RealName="" then
        response.redirect "error.htm"
        response.end
    else
        sql="select * from admin_info where admin_info_id="&id
        Set rs= Server.CreateObject("ADODB.Recordset")
        rs.open sql,conn,1,3
        rs("admin_info_RealName")=admin_info_RealName
        rs("admin_info_flag")    =a
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功修改了一个管理人员信息！","admin_info_list.asp")
    end if
end sub
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理员-管理人员信息-添加</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="admin_info_modi.asp" method="post" name="form1">
<input type="hidden" name="action" value="save">
<input type="hidden" name="admin_info_id" value="<%=id%>"> 
	<tr>
		<td colspan="2" class="header">管理人员-修改</td>
	</tr>
	<tr>
		<td>管理员真实姓名：</td>
		<td><input type="text" name="admin_info_RealName" size="20" value="<%=admin_info_RealName%>"></td>
	</tr>
	<tr>
		<td>登陆用户名：</td>
		<td><font color="#FF0000"><%=admin_info_username%></font>
		<font color="#808080">(注：用户名不可修改)</font></td>
	</tr>
	<tr>
		<td>权限分配：</td>
		<td>
		<table border="1" width="100%" id="table1" cellpadding="4" style="border-collapse: collapse" bordercolor="#CCCCCC">
			<tr>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">基本设置</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">商品管理</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">订单管理</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">会员管理</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">新闻管理</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">留言评论</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">广告管理</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">帮助中心</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">友情链接</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">管理人员</td>
			</tr>
			<tr>
            <%
	        fla=split(admin_info_flag,",")
            for i=0 to ubound(fla)
            %>
		       <td class="altbg2" align="center"><input type="checkbox" name="<%=i%>" value="1" <%if fla(i)=1 then response.write "checked" %>></td>
            <%next%>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="提交" name="B1">&nbsp;
		<input type="reset" value="重置" name="B2"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

 
