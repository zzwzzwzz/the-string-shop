<!--#include file="../include/DuoDuoCode.asp"-->
<%
if session("admin_info_UserName")="" or Session("pass")=False then
    response.redirect "Admin_login.asp"
    response.end
end if
%>
 
