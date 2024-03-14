<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=7
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""��������!"");location.href=""help_info_List.asp"";</script>")
  response.end
end if

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select help_info_title, help_info_content from help_info where id="&id
rs.open sql,conn,1,1
help_info_title	= rs(0)
help_info_content = rs(1)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    id				  = my_request("id",1)
    help_info_title   = my_request("help_info_title",0)
    help_info_content = my_request("Content",0)
    ErrMsg=""
    if id="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>������ϢID����Ϊ�գ�</li>"
    end if
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
        sql="select * from help_info where id="&id
        rs.open sql,conn,1,3
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
<title>������Ϣ-�༭</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Help_Info_Modi.asp" method="post" name="form1" onsubmit="return checkdata();">
<input type="hidden" name="action" value="save">
<input type="hidden" name="id" value="<%=id%>"> 
	<tr>
		<td colspan="2" class="header">������Ϣ-�༭</td>
	</tr>
	<tr>
		<td>��Ϣ���⣺</td>
		<td>
		<input type="text" name="help_info_title" size="40" value="<%=help_info_title%>"></td>
	</tr>
	<tr>
		<td>��Ϣ���ݣ�</td>
		<td>
		    <textarea cols=80 rows=20 id="content" name="Content"><%= Server.HTMLEncode(help_info_content) %></textarea>
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