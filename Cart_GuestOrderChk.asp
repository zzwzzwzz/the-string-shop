<%
dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<%
if session("user_info_id")<>"" then
    response.redirect "Cart_Order.asp"
else
    Set rs=Server.CreateObject("ADODB.Recordset")
    sql="select root_option_GuestOrderOnOff from root_option where id=1"
    rs.open sql,conn,1,1
    root_option_GuestOrderOnOff=rs(0)
    rs.close
    set rs=nothing

    if root_option_GuestOrderOnOff=1 then
        response.redirect "Cart_Order.asp"
    else
        response.redirect "Cart_UserChk.asp" 
    end if
end if
%>
 
