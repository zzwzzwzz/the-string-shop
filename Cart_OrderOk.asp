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

<table border="0" align="center" width="1000" cellpadding="6" style="border-left:2px solid #654321; border-right:2px solid #654321; border-top:1px solid #654321; border-bottom:1px solid #654321; border-collapse: collapse" bgcolor="#654321">
	<tr>
		<td>
		<b><span style="font-size: 14px"><font color="#ffffff">订单提交结果：</font></span></b></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF"><span style="font-size: 14px">
		<font color="#000000"><b>您的订单已提交成功了!</b></font></span><font color="#FF6600">
		</font>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" style="line-height: 150%">
		<span style="font-size: 12px">&nbsp;&nbsp;您的订单号是：<span style="font-size: 14px"><font color="#FF6600"><b><%=order_info_No%></b></font></span>
<br>
&nbsp;&nbsp;&nbsp;&nbsp; </b><font color="#654321">(记下此订单号，以便以后查询订单状态用)</font><br>
		&nbsp;&nbsp;您的订单总支付金额是：<span style="font-size: 14px"><font color="#FF6600"><b><%=order_info_AllCost%>元</b></font></span></span>
		<li><span style="font-size: 12px">选择在线支付方式付款的顾客请继续操作以完成在线支付货款。</span></li>
		<li><span style="font-size: 12px">选择银行汇款或邮局汇款的顾客，请及时将款项汇出，以便我们确认后给你发货,谢谢。</span></li>
		</td>
	</tr>
	<tr>
		<td>
		<p align="left"><span style="font-size: 12px">
		<script language="javascript">
		function PrintIt()
		{window.print()}
		</script>
		<input type="button" style="COLOR:black; border:'2'" value="打印" onClick="PrintIt()" >&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" style="COLOR:black; border:'2'" value="返回" onClick="javascript:location.href='/Cart_Order.asp'" >
		</span>
		</td>
	</tr>
</table>

</body>

</html>

</center>