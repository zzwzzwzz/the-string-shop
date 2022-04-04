<%
user_info_id1=session("user_info_id")
if session("user_info_id")<>"" then
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select user_info_mark from user_info where user_info_id="&user_info_id1
	rs.open sql,conn,1,1
	user_info_mark1=rs(0)
	rs.close
	set rs=nothing

    sql="select user_level_Name from user_Level where user_level_markmin<="&user_info_mark1&" and user_level_markmax>="&user_info_mark1&""
  	set rs=conn.execute (sql)
  	user_level_Name1=rs(0)
  	rs.close
  	set rs=nothing
end if

response.write  "<tr><td colspan=2>欢迎您：<b><font color=#FFb5b5>"&session("user_info_UserName")&"</font></b>&nbsp;<a href=User_LoginOut.asp>[退出登录]</a></td></tr><tr><td colspan=2 align=center height=30><a href=User_Index.asp>帐户首页</a> &nbsp;|&nbsp;"&_ 
				"<a href=User_Personal.asp>修改地址</a> &nbsp;|&nbsp;"&_
				"<a href=User_PassWord.asp>修改密码</a> &nbsp;|&nbsp;"&_
				"<a href=User_OrderList.asp>我的订单</a> &nbsp;|&nbsp;"&_
				"<a href=User_fav.asp>我的收藏</a>"&_
				"</td></tr>"&_
				"<tr><td colspan=2 height=10></td></tr>"
%>
