<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=3
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("user_info_id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""参数错误!"");location.href=""user_info_List.asp"";</script>")
  response.end
end if

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select user_info_RealName,user_info_email,user_info_mobile,user_info_address,user_info_zip,user_info_email,user_info_lastlogintime,user_info_loginNums,user_info_states,user_info_RegTime,user_info_UserName from user_info where user_info_id="&id
rs.open sql,conn,1,1
user_info_RealName=rs(0)
user_info_email=rs(1)
user_info_mobile=rs(2)
user_info_address=rs(3)
user_info_zip=rs(4)
user_info_email=rs(5)
user_info_LastLoginTime=rs(6)
user_info_LoginNums=rs(7)
user_info_states=rs(8)
user_info_RegTime=rs(9)
user_info_UserName=rs(10)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    id                 =my_request("user_info_id",1)
    user_info_states   =my_request("user_info_states",1)
    user_info_PassWord =my_request("user_info_PassWord",0)
    user_info_RealName =my_request("user_info_RealName",0)
    user_info_email    =my_request("user_info_email",0)
    user_info_mobile   =my_request("user_info_mobile",0)
    user_info_address  =my_request("user_info_address",0)
    user_info_zip      =my_request("user_info_zip",0)
    if id="" or user_info_states="" or user_info_RealName="" then
        call error()
    else
        if user_info_PassWord<>"" then
       	    user_info_PassWord=md5(user_info_PassWord,32)
        end if
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from user_info where user_info_id="&id
        rs.open sql,conn,1,3
        rs("user_info_states")  =user_info_states
        if user_info_PassWord<>"" then
        rs("user_info_PassWord")=user_info_PassWord
        end if
        rs("user_info_RealName")=user_info_RealName
        rs("user_info_email")   =user_info_email
        rs("user_info_mobile")  =user_info_mobile
        rs("user_info_address") =user_info_address
        rs("user_info_zip")     =user_info_zip
        rs.update
        rs.close
        set rs=nothing
        call ok("恭喜，您已成功更新了会员信息！","user_info_modi.asp?user_info_id="&id&"")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员-会员信息-查看/编辑</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="user_info_modi.asp" method="post">
<input type="hidden" name="action" value="save"> 
<input type="hidden" name=user_info_id value="<%=id%>">
	<tr>
		<td colspan="2" class="header">查看/编辑会员资料</td>
	</tr>
	<tr>
		<td>用户名：</td>
		<td><%=user_info_UserName%></td>
	</tr>
<tr>
		<td>注册时间：</td>
		<td><%=user_info_RegTime%></td>
	</tr>
<tr>
		<td>上次登陆时间：</td>
		<td><%=user_info_LastLoginTime%></td>
	</tr>
<tr>
		<td>登陆次数：</td>
		<td><%=user_info_LoginNums%> 次</td>
	</tr>
<tr>
		<td>真实姓名：</td>
		<td>
		<input type="text" name="user_info_RealName" size="20" value="<%=user_info_RealName%>"></td>
	</tr>
<tr>
		<td>电子邮箱：</td>
		<td>
		<input type="text" name="user_info_email" size="20" value="<%=user_info_email%>"></td>
	</tr>
<tr>
		<td>联系地址：</td>
		<td>
		<input type="text" name="user_info_address" size="20" value="<%=user_info_address%>"></td>
	</tr>
	<tr>
		<td>邮政编码：</td>
		<td>
		<input type="text" name="user_info_zip" size="20" value="<%=user_info_zip%>"></td>
	</tr>
	<tr>
		<td>联系电话：</td>
		<td>
		<input type="text" name="user_info_mobile" size="20" value="<%=user_info_mobile%>"></td>
	</tr>
	<tr>
		<td>会员状态：</td>
		<td>
		<input type="radio" value="0" name="user_info_states" <%if user_info_states=0 then response.write "checked"%> checked> 正常/通过审核 
		<input type="radio" value="1" name="user_info_states" <%if user_info_states=1 then response.write "checked"%>>  锁定/未审核</td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value=" 提交修改 " name="Submit1">&nbsp;&nbsp;&nbsp;&nbsp; 
		   <input type="button" value=" 返回列表 " name="action1" onClick="window.location='user_info_list.asp'">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
 