<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
call up("级别与积分规则","级别与积分规则","级别与积分规则")

response.write "<tr><td><table width=100% border=1 cellpadding=4 style='border-collapse: collapse' bordercolor=#C0C0C0><tr align=center bgcolor=#E4E4E4><td>会员级别名称</td><td>要求积分下限</td><td>要求积分上限</td><td>享受折扣优惠(不含配送费)(100表示不打折)</td></tr>"
set rs=server.createobject("adodb.recordset")
sql="select id,user_level_name,user_level_markmin,user_level_markmax,user_level_rebate from user_level order by id desc"
rs.open sql,conn,1,1
if rs.eof then 
    response.write "<tr><td colspan=4 align=center><font color=red>目前暂无会员级别信息,请<a href=?action=level_add>点此添加!</a></font></td></tr>"
else
    set id=rs(0)
    set user_level_name=rs(1)
    set user_level_markmin=rs(2)
    set user_level_markmax=rs(3)
    set user_level_rebate=rs(4)

    while not rs.eof
	response.write  "<tr>"&_
					"<td>"&user_level_name&"</td>"&_
					"<td>"&user_level_markmin&"</td>"&_
					"<td>"&user_level_markmax&"</td>"&_
					"<td><b><font color=#FF3300>"&user_level_rebate&"</font></b>折优惠</td>"&_
					"</tr>"
	rs.movenext
	wend
end if
rs.close
set rs=nothing
response.write "</table></tr></tr>"
call down()
%></center>