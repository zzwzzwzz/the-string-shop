<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=7
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    help_info_title   = my_request("help_info_title",0)
    help_info_content = my_request("Content",0)
    ErrMsg=""
    if help_info_title="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��Ϣ���ⲻ��Ϊ�գ�</li>"
    end if
    if help_info_content="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��Ϣ���ݲ���Ϊ�գ�</li>"
    end if
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from help_info"
        rs.open sql,conn,1,3
        rs.addnew
        rs("help_info_title")   = help_info_title
        rs("help_info_content") = help_info_content
        rs.update
        rs.close
        set rs=nothing
        call ok("���ѳɹ�������һ��������Ϣ��","help_info_List.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ϣ-����</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="help_info_add.asp" method="post" name="form1">
<input type="hidden" name="action" value="save">
	<tr>
		<td colspan="2" class="header">������Ϣ-����</td>
	</tr>
	<tr>
		<td>��Ϣ���⣺</td>
		<td><input type="text" name="help_info_title" size="40"></td>
	</tr>
	<tr>
		<td>��Ϣ���ݣ�</td>
		<td>
		    <textarea cols=60 rows=20 id="content" name="Content"></textarea>
        </td>
    </tr>
	<tr>
		<td>��</td>
		<td>
		    <input type="submit" value="  ��  ��  " name="B1">&nbsp; 
		    <input type="reset" value="  ��  ��  " name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>
</html>