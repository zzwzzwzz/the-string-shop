<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=9
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/md5.asp"-->
<%
action=my_request("action",0)
if action="save" then
    call Save()
end if

//管理人员-添加保存
sub Save() 
    admin_info_RealName =my_request("admin_info_RealName",0) 
    admin_info_UserName =my_request("admin_info_UserName",0) 
    admin_info_PassWord =my_request("admin_info_PassWord",0) 
    admin_info_PassWord2=my_request("admin_info_PassWord2",0) 
    for i=0 to 9
        b=request(i)
        if b="" then b=0
        a=a&","&b
    next
    a=right(replace(a," ",""),len(replace(a," ",""))-1)

    if admin_info_RealName="" or admin_info_UserName="" or admin_info_PassWord="" or admin_info_PassWord<>admin_info_PassWord2 then
        response.redirect "error.htm"
        response.end
    else
        sql="select * from admin_info where admin_info_UserName='"&admin_info_UserName&"'"
        Set rs= Server.CreateObject("ADODB.Recordset")
        rs.open sql,conn,1,3
        if not rs.eof then
            response.write"<SCRIPT language=JavaScript>alert('此用户名已被占用了，请重新取一个！');"
            response.write"javascript:history.go(-1)</SCRIPT>"
            Response.end
        else
            rs.addnew
            rs("admin_info_RealName")=admin_info_RealName
            rs("admin_info_UserName")=admin_info_UserName
            rs("admin_info_PassWord")=md5(admin_info_PassWord,32)
            rs("admin_info_flag")=a
            rs.update
        end if
        rs.close
        set rs=nothing
        call ok("您已成功添加了一个新管理人员信息！","admin_info_list.asp")
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
<form action="admin_info_add.asp" method="post" name="form1">
<input type="hidden" name="action" value="save">
	<tr>
		<td colspan="2" class="header">管理人员-添加</td>
	</tr>
	<tr>
		<td>管理员真实姓名：</td>
		<td><input type="text" name="admin_info_RealName" size="20"></td>
	</tr>
	<tr>
		<td>登陆用户名：</td>
		<td><input type="text" name="admin_info_UserName" size="20"></td>
	</tr>
	<tr>
		<td>登陆密码：</td>
		<td><input type="password" name="admin_info_PassWord" size="20"></td>
	</tr>
	<tr>
		<td>再输一次密码：</td>
		<td><input type="password" name="admin_info_PassWord2" size="20"></td>
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
				<td bgcolor="#EFEFEF" class="altbg1" align="center">文章管理</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">留言管理</td>
				<td bgcolor="#EFEFEF" class="altbg1" align="center">权限管理</td>
			</tr>
			<tr>
	         <%for i=0 to 6%>
		        <td align="center"><input type="checkbox" name="<%=i%>" value="1"></td>
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

 
