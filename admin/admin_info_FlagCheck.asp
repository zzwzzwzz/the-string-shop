<%
set rs=server.createobject("adodb.recordset")
sql="select admin_info_flag from admin_info where admin_info_UserName='"&session("admin_info_UserName")&"'"
rs.open sql,conn,1,1
k=split(rs(0),",")
rs.close
set rs=nothing

if k(nownum)<>1 then
    response.write "<script>alert('操作权限出错，您没有权限操作此功能,请向超级管理员申请');history.go(-1);</Script>"
    Response.End 
end if
%>
 
