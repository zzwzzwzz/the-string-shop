<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=0
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<html>
<head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background-color: #FFFFFF">
<%
set rs=server.createobject("adodb.recordset")
sql="select root_info_aboutus from root_info where id=1"
rs.open sql,conn,1,1
If Not rs.Eof Then      
    Content=rs(0)
end if
Response.Write Content
rs.close
set rs=nothing
%>
</body>

</html>

