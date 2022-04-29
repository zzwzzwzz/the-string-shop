<!--#include file="admin_check.asp"-->
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_info_skin from root_info where id=1"
rs.open sql,conn,1,1
root_info_skin    =rs(0)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    root_info_skin            =my_request("root_info_skin",0)
    if root_info_skin="" then
        response.redirect "error.htm"
        response.end
    else
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from root_info where id=1"
        rs.open sql,conn,1,3
        rs("root_info_skin")            =root_info_skin
        rs.update
        rs.close
        set rs=nothing

        call ok("您已成功保存模板设置！","root_model_set.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>基本-网站模板-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="Root_Model_Set.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td class="header">网站模板- 设置</td>
	</tr>
	<tr>
		<td>
		      <table border="0" width="100%" id="table1" cellspacing="1" cellpadding="4" style="border-collapse: collapse">
			      <tr>
				      <td align="center">
				      <img border="0" src="model/temp1.jpg" width="188" height="263"></td>
				      <td align="center">
						<img border="0" src="model/temp2.jpg" width="188" height="266"></td>
			      </tr>
			      <tr>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="default" name="root_info_skin" <%if root_info_skin="default" then response.write "checked"%>>时尚蓝<span class="tit">(默认)</span></td>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="2" name="root_info_skin" <%if root_info_skin="2" then response.write "checked"%>>银灰色</td>
			      </tr>
			      <tr>
				      <td align="center">
				      <img border="0" src="model/temp3.jpg" width="188" height="262"></td>
				      <td align="center">
						<img border="0" src="model/temp4.jpg" width="188" height="266"></td>
			      </tr>
					<tr>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="9" name="root_info_skin" <%if root_info_skin="9" then response.write "checked"%>>经典黑</td>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="1" name="root_info_skin" <%if root_info_skin="10" then response.write "checked"%>>水晶朱</td>
			      </tr>
			      <tr>
				      <td align="center">
				      <img border="0" src="model/temp5.jpg" width="188" height="260"></td>
				      <td align="center">
						<img border="0" src="model/temp6.jpg" width="188" height="258"></td>
			      </tr>
					<tr>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="5" name="root_info_skin" <%if root_info_skin="5" then response.write "checked"%>>花瓣紫</td>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="6" name="root_info_skin" <%if root_info_skin="6" then response.write "checked"%>>3d紫</td>
			      </tr>
			      <tr>
				      <td align="center">
				      <img border="0" src="model/temp7.jpg" width="188" height="266"></td>
				      <td align="center">
						<img border="0" src="model/temp8.jpg" width="188" height="266"></td>
			      </tr>
					<tr>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="3" name="root_info_skin" <%if root_info_skin="3" then response.write "checked"%>>淡雅红</td>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="4" name="root_info_skin" <%if root_info_skin="4" then response.write "checked"%>>浅草绿</td>
			      </tr>
			      <tr>
				      <td align="center">
				      11</td>
				      <td align="center">
						12</td>
			      </tr>
			      <tr>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="11" name="root_info_skin" <%if root_info_skin="11" then response.write "checked"%>>古典</td>
				      <td align="center" bgcolor="#EFEFEF">
				      <input type="radio" value="12" name="root_info_skin" <%if root_info_skin="12" then response.write "checked"%>>浅草绿</td>
			      </tr>
			      </table>
		      </td>
	</tr>
	
	<tr>
		<td>
		<p align="center"><input type="submit" value=" 保存设置  " name="B1">&nbsp;&nbsp;&nbsp;
		<input type="reset" value="重置" name="B2"></td>
	</tr>
	</form>
 </tbody>
</table>
<br>
</body>

</html>