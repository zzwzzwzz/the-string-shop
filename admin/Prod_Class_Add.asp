<!--#include file="admin_check.asp"-->
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    prod_class_name=my_request("prod_class_name",0)
    
    ErrMsg=""
    if prod_class_name="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>Ŀ¼���Ʋ���Ϊ�գ�</li>"
    end if
    
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from prod_class"
        rs.open sql,conn,1,3
        rs.addnew
        rs("prod_class_name")=prod_class_name
        rs.update
        rs.close
        set rs=nothing
        call ok("���ѳɹ������һ��Ŀ¼��","prod_class_list.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ʒ������</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="prod_Class_Add.asp" method="post" name="form1">
<input type="hidden" name="action" value="save">
    <tr>
		<td colspan="2" class="title">��Ʒ������</td>
	</tr>
	<tr>
		<td>������ƣ�</td>
		<td><input type="text" name="prod_class_name" size="20"></td>
	</tr>
	<tr>
		<td>��</td>
		<td><input type="submit" value="�ύ" name="B1">&nbsp;&nbsp;&nbsp;
			<input type="reset" value="����" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>

