<%dim dbpath
dbpath="../"
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
sql="select count(order_info_id) as num0 from order_info where order_info_CheckStates=0 and order_info_recycle=0"
set rs=conn.execute (sql)
num0=rs("num0")
rs.close
set rs=nothing		
		
sql="select count(order_info_id) as num1 from order_info where order_info_CheckStates=1 and order_info_recycle=0"
set rs=conn.execute (sql)
num1=rs("num1")
rs.close
set rs=nothing	
		
sql="select count(order_info_id) as num2 from order_info where order_info_CheckStates=2 and order_info_recycle=0"
set rs=conn.execute (sql)
num2=rs("num2")
rs.close
set rs=nothing	

sql="select count(order_info_id) as num3 from order_info where order_info_CheckStates=3 and order_info_recycle=0"
set rs=conn.execute (sql)
num3=rs("num3")
rs.close
set rs=nothing		
		
sql="select count(order_info_id) as num4 from order_info where order_info_CheckStates=4 and order_info_recycle=0"
set rs=conn.execute (sql)
num4=rs("num4")
rs.close
set rs=nothing	
		
sql="select count(order_info_id) as num5 from order_info where order_info_CheckStates=5 and order_info_recycle=0"
set rs=conn.execute (sql)
num5=rs("num5")
rs.close
set rs=nothing	

sql="select count(order_info_id) as num6 from order_info where order_info_CheckStates=6 and order_info_recycle=0"
set rs=conn.execute (sql)
num6=rs("num6")
rs.close
set rs=nothing

sql="select sum(order_buy_ProdPrice*order_buy_ProdNum) as sumsell from order_buy"
set rs=conn.execute (sql)
sumsell=rs("sumsell")
rs.close
set rs=nothing

sql="select sum(order_buy_ProdNum) as sumnum from order_buy"
set rs=conn.execute (sql)
sumnum=rs("sumnum")
rs.close
set rs=nothing

sql="select count(id) as pnum from product_info"
set rs=conn.execute (sql)
pnum=rs("pnum")
rs.close
set rs=nothing

sql="select count(prod_BigClass_id) as bnum from prod_BigClass"
set rs=conn.execute (sql)
bnum=rs("bnum")
rs.close
set rs=nothing

sql="select count(prod_SmallClass_id) as snum from prod_SmallClass"
set rs=conn.execute (sql)
snum=rs("snum")
rs.close
set rs=nothing

sql="select count(prod_review_id) as prnum from prod_review"
set rs=conn.execute (sql)
prnum=rs("prnum")
rs.close
set rs=nothing

sql="select count(guest_info_id) as gnum from guest_info"
set rs=conn.execute (sql)
gnum=rs("gnum")
rs.close
set rs=nothing

sql="select count(user_info_id) as unum from user_info"
set rs=conn.execute (sql)
unum=rs("unum")
rs.close
set rs=nothing
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨-��ҳ</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td class="header">��̨��ҳ</td>
	</tr>
	<tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<tr>
		<td class="altbg1">��Ϣ����</td>
	</tr>
	<tr>
		<td>
		<p class="p2">
				<a href="order_info_list.asp">&nbsp;&nbsp;����&nbsp;&nbsp;<%=num0%>&nbsp; ���¶����ȴ�����</a></td>
	</tr>
	<tr>
		<td class="altbg1">��Ϣͳ��</td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" id="table1" cellpadding="4" style="border-collapse: collapse">
			<tr>
				<td valign="top" style="border-bottom: 1px solid #E4E4E4"><b>������Ϣ��
		    </b>
		   </td>
				<td style="border-bottom: 1px solid #E4E4E4"> <li>��Ա����ȡ������&nbsp;&nbsp; ��&nbsp;&nbsp;<%=num1%>&nbsp; ��</li>
		    <li>����Աȡ������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��&nbsp;&nbsp;<%=num2%>&nbsp; ��</li>
		    <li>��ȷ�ϣ����������&nbsp;&nbsp;<%=num3%>&nbsp; ��</li>
		    <li>�ѷ��������ջ�������&nbsp;&nbsp;<%=num4%>&nbsp; ��</li>
		    <li>����֧����ɶ���&nbsp;&nbsp; ��&nbsp;&nbsp;<%=num5%>&nbsp; ��</li>
		    <li>��������ɶ���&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��&nbsp;&nbsp;<%=num6%>&nbsp; ��</td>
			</tr>
			<tr>
				<td style="border-bottom: 1px solid #E4E4E4">
			<b>����״����</b></td>
				<td style="border-bottom: 1px solid #E4E4E4"><li>�����������������<%=sumnum%></li> 
				<li>���۶<%=sumsell%> RMB</li>
			</td>
			</tr>
			<tr>
				<td style="border-bottom: 1px solid #E4E4E4"><b>��Ʒͳ�ƣ� </b></td>
				<td style="border-bottom: 1px solid #E4E4E4"><li>�����<%=bnum%> </li> <li>С���<%=snum%> </li> <li>��Ʒ������<%=pnum%> </li>
				<li>��Ʒ���ۣ�<%=prnum%> ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </li></td>
			</tr>
			<tr>
				<td style="border-bottom: 1px solid #E4E4E4"><b>������Ϣ��</b></td>
				<td style="border-bottom: 1px solid #E4E4E4"><%=gnum%> ��</td>
			</tr>
			<tr>
				<td><b>��Ա��Ϣ��</b></td>
				<td><%=unum%> ��</td>
			</tr>
		</table>
		</td>
	</tr>
</tbody>
</table>

</body>

</html>