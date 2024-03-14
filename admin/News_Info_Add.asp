<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=4
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
    news_info_title = my_request("news_info_title",0)
    news_info_content = my_request("Content",0)
    ErrMsg=""
    if news_info_title="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>������Ϣ���ⲻ��Ϊ�գ�</li>"
    end if
    if news_info_content="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>������Ϣ���ݲ���Ϊ�գ�</li>"
    end if
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from news_info"
        rs.open sql,conn,1,3
        rs.addnew
        rs("news_info_title")   = news_info_title
        rs("news_info_content") = news_info_content
        rs("news_info_addtime") = now()
        rs.update
        rs.close
        set rs=nothing
        call ok("���ѳɹ�������һ������������Ϣ��","news_info_List.asp")
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
<form action="news_info_add.asp" method="post" name="form1" onsubmit="return checkdata();">
<input type="hidden" name="action" value="save">
	<tr>
		<td colspan="2" class="header">������Ϣ-����</td>
	</tr>
	<tr>
		<td>���±��⣺</td>
		<td><input type="text" name="news_info_title" size="40"></td>
	</tr>
	<tr>
		<td>�������ݣ�</td>
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