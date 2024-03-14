<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=5
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("guest_info_id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""��������!"");location.href=""GB_Info_List.asp"";</script>")
  response.end
end if

sql="select * from guest_info where guest_info_id="&id
set rs=conn.execute (sql)
guest_info_name=rs("guest_info_name")
guest_info_email=rs("guest_info_email")
guest_info_detail=rs("guest_info_detail")
guest_info_time=rs("guest_info_time")
guest_info_BackDetail=rs("guest_info_BackDetail")
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    id                   =my_request("guest_info_id",1)
    guest_info_email     =my_request("guest_info_email",0)
    guest_info_detail    =my_request("guest_info_detail",0)
    guest_info_BackDetail=my_request("guest_info_BackDetail",0)

    if id="" or guest_info_detail="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from guest_info where guest_info_id="&id
        rs.open sql,conn,1,3
        rs("guest_info_email")     =guest_info_email
        rs("guest_info_detail")    =guest_info_detail
        rs("guest_info_BackDetail")=guest_info_BackDetail
        rs("guest_info_BackTime")  =now()
        rs.update
        rs.close
        set rs=nothing
        call ok("���ѳɹ��ظ�/������һ��������Ϣ��","GB_Info_List.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ϣ-�ظ�</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="gb_info_back.asp" method="post">
<input type="hidden" name="action" value="save"> 
<input type="hidden" name="guest_info_id" value="<%=id%>"> 
	<tr>
		<td colspan="2" class="header">������Ϣ-�ظ�</td>
	</tr>
	<tr>
		<td>����ʱ�䣺</td>
		<td><%=guest_info_time%></td>
	</tr>
	<tr>
		<td>������������</td>
		<td><%=guest_info_name%></td>
	</tr>
	<tr>
		<td>�����ʼ���</td>
		<td>
		<input type="text" name="guest_info_email" size="30" value="<%=guest_info_email%>"></td>
	</tr>
		<tr>
		<td>�������ݣ�</td>
		<td><textarea rows="8" name="guest_info_detail" cols="60"><%=guest_info_detail%></textarea></td>
	</tr>
	<tr>
		<td>�ظ����ݣ�</td>
		<td><textarea rows="8" name="guest_info_BackDetail" cols="60"><%=guest_info_BackDetail%></textarea></td>
	</tr>
	<tr>
		<td>��</td>
		<td><input type="submit" value=" ��  �� " name="Submit1">&nbsp; 
		   <input type="reset" value="����" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>
</html>