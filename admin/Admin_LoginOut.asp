<!--#include file="Admin_check.asp"-->
<!--#include file="../include/DuoDuoCode.asp"-->
<%dim dbpath
dbpath="../"
%>
<!--#include file="../Conn.asp"-->
<%
session("admin_info_UserName")=""
session("admin_info_RealName")=""
session("pass")=""

//定期检查清理收藏夹(保留一个月)
conn.execute ("delete from [prod_favorite] where DateDiff('d', prod_favorite_time, now)>30")

response.redirect "../admin/Admin_Login.asp" 
%>
 
