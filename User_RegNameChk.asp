<%dim dbpath,action
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp"-->
<%
response.write "<title>�û������</title>"
response.write "<span style='font-size: 12px'>"
userName=trim(request("username"))
if userName="" or len(userName)<4 then
    response.write "<li><font color=#ff0000>�û�������Ϊ�ջ�С��4λ����</font></li>"
    response.end
else
    set rs=server.createobject("adodb.recordset")
    sql="select * from User_info where User_info_UserName='"&username&"'" 
    rs.open sql,conn,1,1
    if rs.bof and rs.eof then
	    response.write "<center><li><font color=#ff0000>��ϲ�������û�������ʹ�ã�</font></li>"
    else
	    response.write "<li><font color=#ff0000>�Բ��𣬴��û����ѱ�ռ�ã�������ȡһ����</font></li>"
	    response.end
    end if
    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
end if
response.write "</span>"
%>
