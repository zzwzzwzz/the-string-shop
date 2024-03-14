<center>
<!--#include file="User_Chk.asp"-->
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
'ȡ������
id=session("user_info_id")

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select user_info_RealName,user_info_email,user_info_mobile,user_info_address,user_info_zip from user_info where user_info_id="&id
rs.open sql,conn,1,1
user_info_RealName=rs(0)
user_info_email=rs(1)
user_info_mobile=rs(2)
user_info_address=rs(3)
user_info_zip=rs(4)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call User_PersonalModiSave()
end if

call up("�ҵ��ʻ���ҳ","�ҵ��ʻ���ҳ","�ҵ��ʻ���ҳ")
%>
<!--#include file="User_Menu.asp"-->
<%
response.write  "<tr><td colspan=2><b>��������:</b></td></tr>"&_
				"<tr><td>���� :</td><td>"&user_info_realname&"</td></tr>"&_
				"<tr><td>Email :</td><td>"&user_info_Email&"</td></tr>"&_
				"<tr><td>�ջ���ַ:</td><td>"&user_info_address&"</td></tr>"&_
				"<tr><td>��������:</td><td>"&user_info_zip&"</td></tr>"&_
				"<tr><td>��ϵ�绰:</td><td>"&user_info_mobile&"</td></tr>"&_
				"</form>"
call down()
%>
</center>