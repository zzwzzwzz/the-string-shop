<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
dim root_info_ICP,root_info_tel,root_info_email,root_info_QQ,root_info_MSN,root_info_WangWang,root_info_address,root_info_zip,root_info_fax
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_info_ICP,root_info_tel,root_info_email,root_info_QQ,root_info_MSN,root_info_WangWang,root_info_address,root_info_zip,root_info_fax from root_info where id=1"
rs.open sql,conn,1,1
root_info_ICP             =rs(0)
root_info_tel             =rs(1)
root_info_email           =rs(2)
root_info_QQ              =rs(3)
root_info_MSN             =rs(4)
root_info_WangWang        =rs(5)
root_info_address 		  =rs(6)
root_info_zip			  =rs(7)
root_info_fax			  =rs(8)
rs.close
set rs=nothing

call up("联系我们","联系我们","联系我们")

				if root_info_tel<>"" then
response.write "<tr><td><b>电话：</b></td><td>"&root_info_tel&"</td></tr>"
				end if
				if root_info_email<>"" then
response.write "<tr><td><b>Email：</b></td><td>"&root_info_email&"</td></tr>"
				end if
				if root_info_address<>"" then
response.write "<tr><td><b>地址：</b></td><td>"&root_info_address&"</td></tr>"
				end if
				if root_info_zip<>"" then
response.write "<tr><td><b>邮编：</b></td><td>"&root_info_zip&"</td></tr>"
				end if
call down()
%>
</center>