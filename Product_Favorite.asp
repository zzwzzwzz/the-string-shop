<center><%dim nowplace
nowplace="add_fav"
id=my_request("id",1)
%>
<!--#include file="User_Chk.asp"-->
<%dim dbpath,id,ErrMsg,FoundErr
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp" -->
<%
ErrMsg=""
if id="" or isnull(id) or IsNumeric(id)=False then
    FoundErr=True
	ErrMsg=ErrMsg & "<li>��������</li>"
end if

if FoundErr<>True then
    sql="select * from product_info where id="&id
    set rs=conn.execute (sql)
    if (rs.eof and rs.bof) then
        response.write"<SCRIPT language=JavaScript>alert('�Բ�����վ��û�д���Ʒ��Ϣ��');"
        response.write"javascript:history.go(-1)</SCRIPT>"
        response.end
    end if
    rs.close
    set rs=nothing

    uid=session("user_info_id")
    set rs=server.CreateObject("adodb.recordset")
    sql="select * from prod_favorite where prod_favorite_pid="&id&" and prod_favorite_uid="&uid
    rs.open sql,conn,1,3
    if not (rs.eof and rs.bof) then
        response.write"<SCRIPT language=JavaScript>alert('�Բ������Ѿ��ղع�����Ʒ�������ظ��ղأ�');"
        response.write"javascript:history.go(-1)</SCRIPT>"
        response.end
    else
        rs.addnew
        rs("prod_favorite_pid")=id
        rs("prod_favorite_uid")=uid
        rs("prod_favorite_time")=datevalue(now())
        rs.update
    end if
    rs.close
    set rs=nothing

    response.write"<SCRIPT LANGUAGE=javascript>alert('����Ʒ�ѳɹ���������ղؼУ�');location.href='product_detail.asp?id="&id&"';</SCRIPT>"
    response.end
else
    call WriteErrMsg(WriteErrMsg)
end if
%>

</center>