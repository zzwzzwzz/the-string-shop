<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp" -->
<%
dim order_info_No,order_info_AllCost
order_info_No     =my_request("order_info_No",0)
order_info_AllCost=my_request("order_info_AllCost",0)
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>订购成功！</title>
</head>

<body>

<table border="0" width="100%" cellpadding="6" style="border-left:2px solid #FFE8D9; border-right:2px solid #FFE8D9; border-top:1px solid #FF9900; border-bottom:1px solid #FF9900; border-collapse: collapse" bgcolor="#FFE8D9">
	<tr>
		<td>
		<b><span style="font-size: 14px">订单提交结果：</span></b></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF"><span style="font-size: 14px">
		<font color="#009900"><b>您的订单已提交成功了!</b></font></span><font color="#009900">
		</font>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" style="line-height: 150%">
		<span style="font-size: 12px">&nbsp;&nbsp;您的订单号是：<span style="font-size: 14px"><font color="#FF6600"><b><%=order_info_No%></b></font></span>
<br>
&nbsp;&nbsp;&nbsp;&nbsp; </b><font color="#FF9900">(记下此订单号，以便以后查询订单状态用)</font><br>
		&nbsp;&nbsp;您的订单总支付金额是：<span style="font-size: 14px"><font color="#FF6600"><b><%=order_info_AllCost%>元</b></font></span></span>
		<li><span style="font-size: 12px">选择在线支付方式付款的顾客请继续操作以完成在线支付货款。</span></li>
		<li><span style="font-size: 12px">选择银行汇款或邮局汇款的顾客，请及时将款项汇出，以便我们确认后给你发货,谢谢。</span></li>
		</td>
	</tr>
<%
set rs=server.createobject("adodb.recordset")
sql="select order_info_pay from order_info where order_info_no='"&order_info_No&"'"
rs.open sql,conn,1,1
order_info_pay=rs(0)
rs.close
set rs=nothing
if order_info_pay=6 or order_info_pay=7 then
	set rs1=server.createobject("adodb.recordset")
	sql1="select root_info_remit from root_info where id=1"
	rs1.open sql1,conn,1,1
	If Not rs1.Eof Then      
    	order_info_remit=rs1(0)
	end if
	rs1.close
	set rs1=nothing
	
	response.write "<tr><td><span style='font-size: 14px'><font color=red><b>汇款说明:</b></font></span></td></tr>"
	response.write "<tr><td style='table-layout:fixed;word-break:break-all;line-height: 150%' bgcolor=#FFFFFF>"&order_info_remit&"</td></tr>"
end if
%>

	<tr>
		<td>
		<p align="center"><span style="font-size: 12px">
		<script language="javascript">
		function PrintIt()
		{    window.print()}
		</script>
		<input type="button" style="COLOR:green; border:'2'"value="打印" onClick="PrintIt()" onMouseOver="this.value='打印该页'" onMouseOut="this .value='打印该页'">&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:window.close()">关闭窗口</a></span></td>
	</tr>
</table>

</body>

</html>

</center>