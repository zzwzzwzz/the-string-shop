<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=2
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
select case action
   case "count_time"   '���ڼ�ͳ��
       time_DayFrom=my_request("time_DayFrom",0)
       time_DayTo  =my_request("time_DayTo",0)
       DayFrom=cdate(time_DayFrom)
       DayTo  =cdate(time_DayTo)
       DayTo  =DateAdd("d",1,DayTo)
       DayFrom="#"&DayFrom&"#"
       DayTo="#"&DayTo&"#"
       call count_time()  
   
   case "count_day"    '����ͳ��
       day_day=my_request("day_day",0)
       call count_day()
       
   case "count_month"  '����ͳ��
       month_year =my_request("month_year",0)
       month_month=my_request("month_month",0)
       a=month_year&"-"&month_month
       a=cmonth(a)
       call count_month() 
       
   case "count_season"  '������ͳ��
       season_year  =my_request("season_year",0)
       season_season=my_request("season_season",0)
       select case season_season
           case 1
               a1="#"&season_year&"-1-1#"
               a2="#"&season_year&"-4-1#"
               a1=cmonth(a1)
               a2=cmonth(a2)
           case 2
               a1="#"&season_year&"-4-1#"
               a2="#"&season_year&"-7-1#"
               a1=cmonth(a1)
               a2=cmonth(a2)
           case 3
               a1="#"&season_year&"-7-1#"
               a2="#"&season_year&"-10-1#"           
               a1=cmonth(a1)
               a2=cmonth(a2)
           case 4
               a1="#"&season_year&"-10-1#"
               a2="#"&season_year+1&"-1-1#"
               a1=cmonth(a1)
               a2=cmonth(a2)
       end select
       call count_season() 
       
end select
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����-������Ϣ-����ͳ��</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<%sub count_time()%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">����ͳ��-�ڼ䱨��</td>
	</tr>
	<tr>
		<td>ͳ���ڼ䣺</td>
		<td>�� <%=time_DayFrom%>  ��  <%=time_DayTo%> </td>
	</tr>
	<tr>
		<td>�ڼ�������ϸ��</td>
		<td>��
		<table border="1" width="100%" id="table1" cellpadding="4" style="border-collapse: collapse" bordercolor="#000000">
			<tr>
				<td width="235" bgcolor="#EAEAEA"><b>��Ʒ����</b></td>
				<td width="44" bgcolor="#EAEAEA"><b>����</b></td>
				<td width="82" bgcolor="#EAEAEA"><b>������</b></td>
				<td width="123" bgcolor="#EAEAEA"><b>С��</b></td>
			</tr>
			<%
            set rs=server.createobject("adodb.recordset")
            sql="select order_buy_ProdName,order_buy_ProdPrice,sum(order_buy_ProdNum) as aaa from order_buy where order_buy_BuyTime Between "&DayFrom&" and "&DayTo&" group by order_buy_ProdName,order_buy_ProdPrice"
            rs.open sql,conn,1,1
            if rs.eof then 
                response.write "<tr><td colspan=4 align=center>����û����Ʒ������Ϣ(ֻ�ж������״̬����ͳ�Ʒ�Χ)</td></tr>"
            else
                set order_buy_ProdNum  =rs("aaa")
                set order_buy_ProdName =rs("order_buy_ProdName")
                set order_buy_ProdPrice=rs("order_buy_ProdPrice")
                while not rs.eof
                cost=cost+order_buy_ProdPrice*order_buy_ProdNum
            %>
			<tr>
				<td width="235"><%=order_buy_ProdName%></td>
				<td width="44">��<%=order_buy_ProdPrice%></td>
				<td width="82"><%=order_buy_ProdNum%></td>
				<td width="123">��<%=order_buy_ProdPrice*order_buy_ProdNum%></td>
			</tr>
			<% 	rs.movenext
			    wend
			end if
			rs.close
			set rs=nothing
			%>
		</table>
		</td>
	</tr>
	<tr>
		<td>�ڼ���Ʒ�����۶�(�������ͷ���)��</td>
		<td><b>��<%=cost%></b></td>
	</tr>
</tbody>
</table>
<%
end sub

sub count_day()%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">����ͳ��-�ձ���</td>
	</tr>
	<tr>
		<td>ͳ���գ�</td>
		<td><%=datevalue(day_day)%></td>
	</tr>
	<tr>
		<td>����������ϸ��</td>
		<td>
		<table border="1" width="100%" id="table1" cellpadding="4" style="border-collapse: collapse" bordercolor="#000000">
			<tr>
				<td width="235" bgcolor="#EAEAEA"><b>��Ʒ����</b></td>
				<td width="44" bgcolor="#EAEAEA"><b>����</b></td>
				<td width="82" bgcolor="#EAEAEA"><b>������</b></td>
				<td width="123" bgcolor="#EAEAEA"><b>С��</b></td>
			</tr>
			<%
            set rs=server.createobject("adodb.recordset")
            sql="select order_buy_ProdName,order_buy_ProdPrice,sum(order_buy_ProdNum) as aaa from order_buy where order_buy_BuyTime like '%"&datevalue(day_day)&"%' group by order_buy_ProdName,order_buy_ProdPrice"
            rs.open sql,conn,1,1
            if rs.eof then 
                response.write "<tr><td colspan=4 align=center>����û����Ʒ������Ϣ(ֻ�ж������״̬����ͳ�Ʒ�Χ)</td></tr>"
            else
                set order_buy_ProdNum  =rs("aaa")
                set order_buy_ProdName =rs("order_buy_ProdName")
                set order_buy_ProdPrice=rs("order_buy_ProdPrice")
                while not rs.eof
                cost=cost+order_buy_ProdPrice*order_buy_ProdNum

            %>
			<tr>
				<td width="235"><%=order_buy_ProdName%></td>
				<td width="44">��<%=order_buy_ProdPrice%></td>
				<td width="82"><%=order_buy_ProdNum%></td>
				<td width="123">��<%=order_buy_ProdPrice*order_buy_ProdNum%></td>
			</tr>
			<% 	rs.movenext
			    wend
			end if
			rs.close
			set rs=nothing
			%>
		</table>
		</td>
	</tr>
	<tr>
		<td>������Ʒ�����۶�(�������ͷ���)��</td>
		<td><b>��<%=cost%></b></td>
	</tr>
</tbody>
</table>
<%
end sub

sub count_month()
%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">����ͳ��-�±���</td>
	</tr>
	<tr>
		<td>ͳ���£�</td>
		<td><%=a%></td>
	</tr>
	<tr>
		<td>����������ϸ��</td>
		<td>
		<table border="1" width="100%" id="table1" cellpadding="4" style="border-collapse: collapse" bordercolor="#000000">
			<tr>
				<td width="235" bgcolor="#EAEAEA"><b>��Ʒ����</b></td>
				<td width="44" bgcolor="#EAEAEA"><b>����</b></td>
				<td width="82" bgcolor="#EAEAEA"><b>������</b></td>
				<td width="123" bgcolor="#EAEAEA"><b>С��</b></td>
			</tr>
			<%
            set rs=server.createobject("adodb.recordset")
            sql="select order_buy_ProdName,order_buy_ProdPrice,sum(order_buy_ProdNum) as aaa from order_buy where order_buy_BuyTime like '%"&a&"%' group by order_buy_ProdName,order_buy_ProdPrice"
            rs.open sql,conn,1,1
            if rs.eof then 
                response.write "<tr><td colspan=4 align=center>����û����Ʒ������Ϣ(ֻ�ж������״̬����ͳ�Ʒ�Χ)</td></tr>"
            else
                set order_buy_ProdNum  =rs("aaa")
                set order_buy_ProdName =rs("order_buy_ProdName")
                set order_buy_ProdPrice=rs("order_buy_ProdPrice")
                while not rs.eof
                cost=cost+order_buy_ProdPrice*order_buy_ProdNum

            %>
			<tr>
				<td width="235"><%=order_buy_ProdName%></td>
				<td width="44">��<%=order_buy_ProdPrice%></td>
				<td width="82"><%=order_buy_ProdNum%></td>
				<td width="123">��<%=order_buy_ProdPrice*order_buy_ProdNum%></td>
			</tr>
			<% 	rs.movenext
			    wend
			end if
			rs.close
			set rs=nothing
			%>
		</table>
		
		</td>
	</tr>
	<tr>
		<td>������Ʒ�����۶�(�������ͷ���)��</td>
		<td><b>��<%=cost%></b></td>
	</tr>
</tbody>
</table>
<%
end sub

sub count_season()
%>


<table cellspacing="1" cellpadding="4" width="100%" class="tableborder" id="table3">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">����ͳ��-���ȱ���</td>
	</tr>
	<tr>
		<td>ͳ�Ƽ��ȣ�</td>
		<td><%=season_year%>���<%=season_season%>����</td>
	</tr>
	<tr>
		<td>������������ϸ��</td>
		<td>
		<table border="1" width="100%" id="table4" cellpadding="4" style="border-collapse: collapse" bordercolor="#000000">
			<tr>
				<td width="235" bgcolor="#EAEAEA"><b>��Ʒ����</b></td>
				<td width="44" bgcolor="#EAEAEA"><b>����</b></td>
				<td width="82" bgcolor="#EAEAEA"><b>������</b></td>
				<td width="123" bgcolor="#EAEAEA"><b>С��</b></td>
			</tr>
			<%
            set rs=server.createobject("adodb.recordset")
            sql="select order_buy_ProdName,order_buy_ProdPrice,sum(order_buy_ProdNum) as aaa from order_buy where order_buy_BuyTime between "&a1&" and "&a2&" group by order_buy_ProdName,order_buy_ProdPrice"
            rs.open sql,conn,1,1
            if rs.eof then 
                response.write "<tr><td colspan=4 align=center>����û����Ʒ������Ϣ(ֻ�ж������״̬����ͳ�Ʒ�Χ)</td></tr>"
            else
                set order_buy_ProdNum  =rs("aaa")
                set order_buy_ProdName =rs("order_buy_ProdName")
                set order_buy_ProdPrice=rs("order_buy_ProdPrice")
                while not rs.eof
                cost=cost+order_buy_ProdPrice*order_buy_ProdNum

            %>
			<tr>
				<td width="235"><%=order_buy_ProdName%></td>
				<td width="44">��<%=order_buy_ProdPrice%></td>
				<td width="82"><%=order_buy_ProdNum%></td>
				<td width="123">��<%=order_buy_ProdPrice*order_buy_ProdNum%></td>
			</tr>
			<% 	rs.movenext
			    wend
			end if
			rs.close
			set rs=nothing
			%>
		</table>
		
		</td>
	</tr>
	<tr>
		<td>��������Ʒ�����۶�(�������ͷ���)��</td>
		<td><b>��<%=cost%></b></td>
	</tr>
</tbody>
</table>
<%end sub%>

</body>

</html>
 
