<%
set rs=server.createobject("adodb.recordset")
sql="select admin_info_flag from admin_info where admin_info_UserName='"&session("admin_info_UserName")&"'"
rs.open sql,conn,1,1
k=split(rs(0),",")
rs.close
set rs=nothing

if k(nownum)<>1 then
    response.write "<script>alert('����Ȩ�޳�����û��Ȩ�޲����˹���,���򳬼�����Ա����');history.go(-1);</Script>"
    Response.End 
end if
%>
 
