<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
dim txt_info_title,txt_info_content
id=my_request("id",1)
sql="select help_info_title,help_info_content from help_info where id="&id
set rs=conn.execute (sql)
txt_info_title  =rs(0)
txt_info_content=rs(1)
rs.close
set rs=nothing

call up(help_info_title,"帮助详情","<a href=Help_List.asp>帮助中心</a> &raquo; 帮助详情")

response.write  "<tr><td><h2 align=center>"&txt_info_title&"</h2></td></tr>"&_
				"<tr><td style='table-layout:fixed;word-break:break-all' class=maintxt>"&txt_info_content&"<br></td></tr>"
call down()
%></center>