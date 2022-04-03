<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<html>
<head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
<style>
<!--
body         { background-color: #FFFFFF }
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
action=my_request("action",0)
id=my_request("id",1)
If action="modify" Then
   set rs=server.createobject("adodb.recordset")
   sql="select Product_info_detail from product_info where id="&id
   rs.open sql,conn,1,1
   content=rs(0)
   rs.close
   set rs=nothing
   Response.Write content
End If
%>
</body>

</html>

