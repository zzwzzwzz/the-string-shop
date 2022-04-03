<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=4
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
action=my_request("action",0)
id=my_request("id",1)
If action="modify" Then
   set rs=server.createobject("adodb.recordset")
   sql="select news_info_type,news_info_content from news_info where id="&id
   rs.open sql,conn,1,1
   If Not rs.Eof Then
       if rs(0)=0 then 
           Content=""
       end if
       if rs(0)=1 then        
           Content=rs(1)
       End If
   end if
   Response.Write Content
   rs.close
   set rs=nothing
End If
%>
</body>

</html>

