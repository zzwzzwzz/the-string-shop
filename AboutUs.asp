<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=include/Pages.asp-->
<!--#include file=Sub.asp-->
<%
dim rs,sql,root_info_aboutus
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_info_aboutus from root_info where id=1"
rs.open sql,conn,1,1
root_info_aboutus =rs(0)
rs.close
set rs=nothing
call up ("关于我们","关于我们","关于我们")
response.write "<tr><td><h2 align=center>关于我们</h2></td></tr>"&_
			   "<tr><td style='table-layout:fixed;word-break:break-all'>"&root_info_aboutus&"</td></tr>"
call down()
%>
</center>