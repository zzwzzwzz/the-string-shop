<%
user_info_id1=session("user_info_id")
if session("user_info_id")<>"" then
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select user_info_username from user_info where user_info_id="&user_info_id1
	rs.open sql,conn,1,1
	user_info_username=rs(0)
	rs.close
	set rs=nothing

end if

response.write  "<tr><td colspan=2>��ӭ����<b><font color=#FFb5b5>"&session("user_info_UserName")&"</font></b>&nbsp;<a href=User_LoginOut.asp>[�˳���¼]</a></td></tr><tr><td colspan=2 align=center height=30><a href=User_Index.asp>�ʻ���ҳ</a> &nbsp;|&nbsp;"&_ 
				"<a href=User_Personal.asp>�޸ĵ�ַ</a> &nbsp;|&nbsp;"&_
				"<a href=User_PassWord.asp>�޸�����</a> &nbsp;|&nbsp;"&_
				"<a href=User_OrderList.asp>�ҵĶ���</a> &nbsp;|&nbsp;"&_
				"<a href=User_fav.asp>�ҵ��ղ�</a>"&_
				"</td></tr>"&_
				"<tr><td colspan=2 height=10></td></tr>"
%>
