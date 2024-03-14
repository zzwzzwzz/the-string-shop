<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
call up("帮助中心","帮助中心","帮助中心")

dim rs,sql,id,txt_info_title
set rs=server.createobject("adodb.recordset")
sql="select id,help_info_title from help_info order by id desc"
rs.open sql,conn,1,1
if rs.eof then 
    response.write "暂无帮助信息!"
else
    set id             = rs(0)
    set txt_info_title = rs(1)	
    while not rs.eof
    	response.write "<tr><td><b><a href=Help_Detail.asp?id="&id&">"&txt_info_title&"</b></td></tr>"
    	rs.movenext
    wend
end if
rs.close
set rs=nothing

call down()
%>
</center>