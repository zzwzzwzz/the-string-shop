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
<title>�����ɹ���</title>
</head>

<body>

<table border="0" align="center" width="1000" cellpadding="6" style="border-left:2px solid #654321; border-right:2px solid #654321; border-top:1px solid #654321; border-bottom:1px solid #654321; border-collapse: collapse" bgcolor="#654321">
	<tr>
		<td>
		<b><span style="font-size: 14px"><font color="#ffffff">�����ύ�����</font></span></b></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF"><span style="font-size: 14px">
		<font color="#000000"><b>���Ķ������ύ�ɹ���!</b></font></span><font color="#FF6600">
		</font>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" style="line-height: 150%">
		<span style="font-size: 12px">&nbsp;&nbsp;���Ķ������ǣ�<span style="font-size: 14px"><font color="#FF6600"><b><%=order_info_No%></b></font></span>
<br>
&nbsp;&nbsp;&nbsp;&nbsp; </b><font color="#654321">(���´˶����ţ��Ա��Ժ��ѯ����״̬��)</font><br>
		&nbsp;&nbsp;���Ķ�����֧������ǣ�<span style="font-size: 14px"><font color="#FF6600"><b><%=order_info_AllCost%>Ԫ</b></font></span></span>
		<li><span style="font-size: 12px">ѡ������֧����ʽ����Ĺ˿�������������������֧�����</span></li>
		<li><span style="font-size: 12px">ѡ�����л����ʾֻ��Ĺ˿ͣ��뼰ʱ�����������Ա�����ȷ�Ϻ���㷢��,лл��</span></li>
		</td>
	</tr>
	<tr>
		<td>
		<p align="left"><span style="font-size: 12px">
		<script language="javascript">
		function PrintIt()
		{window.print()}
		</script>
		<input type="button" style="COLOR:black; border:'2'" value="��ӡ" onClick="PrintIt()" >&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" style="COLOR:black; border:'2'" value="����" onClick="javascript:location.href='/Cart_Order.asp'" >
		</span>
		</td>
	</tr>
</table>

</body>

</html>

</center>