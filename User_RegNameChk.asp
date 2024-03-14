<%dim dbpath,action
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp"-->
<%
response.write "<title>用户名检测</title>"
response.write "<span style='font-size: 12px'>"
userName=trim(request("username"))
if userName="" or len(userName)<4 then
    response.write "<li><font color=#ff0000>用户名不能为空或小于4位数！</font></li>"
    response.end
else
    set rs=server.createobject("adodb.recordset")
    sql="select * from User_info where User_info_UserName='"&username&"'" 
    rs.open sql,conn,1,1
    if rs.bof and rs.eof then
	    response.write "<center><li><font color=#ff0000>恭喜您，此用户名可以使用！</font></li>"
    else
	    response.write "<li><font color=#ff0000>对不起，此用户名已被占用，请重新取一个！</font></li>"
	    response.end
    end if
    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
end if
response.write "</span>"
%>
