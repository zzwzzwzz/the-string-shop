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

//���ڼ�������ղؼ�(����һ����)
conn.execute ("delete from [prod_favorite] where DateDiff('d', prod_favorite_time, now)>30")

response.redirect "../admin/Admin_Login.asp" 
%>
