<center><!--#include file="User_Chk.asp"-->
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

call up("�༭�ʻ���Ϣ","�༭�ʻ���Ϣ","�༭�ʻ���Ϣ")
%>
<!--#include file="User_Menu.asp"-->
<%
response.write  "<form name=form1 action=user_Personal.asp method=post>"&_
				"<input type=hidden name=action value=save>"&_
				"<tr><td>�û���:</td><td>"&session("user_info_UserName")&"</td></tr>"&_
				"<tr><td>��  ��:</td><td>****** ( <a href=User_PassWord.asp>&gt;&gt;�޸�����</a> )</td></tr>"&_
				"<tr><td>��  ��:</td><td><input type=text name=user_info_RealName size=30 value="&user_info_RealName&"></td></tr>"&_
				"<tr><td>E-mail:</td><td><input type=text name=user_info_Email size=30 value="&user_info_Email&"></td></tr>"&_
				"<tr><td>�ջ���ַ:</td><td><input type=text name=user_info_address size=30 value="&user_info_address&"></td></tr>"&_
				"<tr><td>��������:</td><td><input type=text name=user_info_zip size=30 value="&user_info_zip&"></td></tr>"&_
				"<tr><td>��ϵ�绰:</td><td><input type=text name=user_info_mobile size=30 value="&user_info_mobile&"></td></tr>"&_
				"<tr><td></td><td><input class=button type=submit value=�ύ�޸�></td></tr>"&_
				"</form>"
call down()
%></center>