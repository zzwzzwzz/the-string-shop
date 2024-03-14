<%dim dbpath
dbpath=""
%>
<!--#include file="conn.asp"-->
<!--#include file="include/MyRequest.asp"-->
<%
item_number=my_request("item_number",0)
response.write "<p>支付已经成功，订单号为："&item_number&"</p>"
conn.execute("Update buyer set zt =1 where ddbh='"&item_number&"'")
conn.close
set conn=nothing
%>

 
