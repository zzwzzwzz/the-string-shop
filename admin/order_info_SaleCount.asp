<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=2
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����-������Ϣ-�ظ�</title>
<link rel="stylesheet" type="text/css" href="style.css">
<SCRIPT language="javascript" src="images/calendar.js"></SCRIPT> 
<SCRIPT language="javascript" src="images/datefunction.js"></SCRIPT>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">����ͳ��</td>
	</tr>
	<tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<!--�ڼ�����ͳ�Ʊ�����  //star -->
    <tr>
		<td colspan="2" class="altbg1">�ڼ�����ͳ�ƣ�</td>
	</tr>
	<form name="form1" method="post" action="order_info_SaleCountDetail.asp" id="form1">
    <input type="hidden" name="action" value="count_time"> 
	<tr>
		<td>ͳ���ڼ�ѡ��</td>
		<td>��<input size="10"  name="Time_DayFrom" id="Time_DayFrom" onclick="event.cancelBubble=true;showCalendar('Time_DayFrom',false,'Time_DayFrom')" value="">&nbsp; 
		�� <input size="10"  name="Time_DayTo" id="Time_DayTo" onclick="event.cancelBubble=true;showCalendar('Time_DayTo',false,'Time_DayTo')" value="">
		<input type="submit" value="�ύ" name="B3"></td>
	</tr>
    </form>
    <!--������ͳ�Ʊ�����  //star -->
    <form name="form2" method="post" action="order_info_SaleCountDetail.asp" id="form2">
    <input type="hidden" name="action" value="count_day"> 	
    <tr>
		<td colspan="2" class="altbg1">������ͳ�ƣ�</td>
		</tr>
    <tr>
		<td>ͳ������ѡ��</td>
		<td><input size="10"  name="day_day" id="day_day" onclick="event.cancelBubble=true;showCalendar('day_day',false,'day_day')" value=""> <input type="submit" value="�ύ" name="B1"></td>
	</tr>
	</form>
	<!--������ͳ�Ʊ�����  //star -->
    <form name="form3" method="post" action="order_info_SaleCountDetail.asp" id="form2">
    <input type="hidden" name="action" value="count_month"> 	
    <tr>
		<td colspan="2" class="altbg1">������ͳ�ƣ�</td>
	</tr>
	<tr>
		<td>ͳ���·�ѡ��</td>
		<td><select size="1" name="month_year">
		<option value="2018" <%if year(now)=2018 then response.write "selected"%>>2018</option>
		<option value="2019" <%if year(now)=2019 then response.write "selected"%>>2019</option>
		<option value="2020" <%if year(now)=2020 then response.write "selected"%>>2020</option>
		<option value="2021" <%if year(now)=2021 then response.write "selected"%>>2021</option>
		<option value="2022" <%if year(now)=2022 then response.write "selected"%>>2022</option>
		<option value="2023" <%if year(now)=2023 then response.write "selected"%>>2023</option>
		<option value="2024" <%if year(now)=2024 then response.write "selected"%>>2024</option>
		<option value="2025" <%if year(now)=2025 then response.write "selected"%>>2025</option>
		<option value="2026" <%if year(now)=2026 then response.write "selected"%>>2026</option>
		<option value="2027" <%if year(now)=2027 then response.write "selected"%>>2027</option>
		<option value="2028" <%if year(now)=2028 then response.write "selected"%>>2028</option>
		</select>�� <select size="1" name="month_month">
		<option value="1">1</option>
		<option value="2">2</option>
		<option value="3">3</option>
		<option value="4">4</option>
		<option value="5">5</option>
		<option value="6">6</option>
		<option value="7">7</option>
		<option value="8">8</option>
		<option value="9">9</option>
		<option value="10">10</option>
		<option value="11">11</option>
		<option value="12">12</option>
		</select>�� <input type="submit" value="�ύ" name="B4"></td>
	</tr>
	</form>
	<!--��������ͳ�Ʊ�����  //star -->
    <tr>
		<td colspan="2" class="altbg1">��������ͳ�ƣ�</td>
	</tr>
	<form name="form4" method="post" action="order_info_SaleCountDetail.asp" id="form4">
    <input type="hidden" name="action" value="count_season"> 
    	<tr>
		<td>ͳ�Ƽ���ѡ��</td>
		<td><select size="1" name="season_year">
		<option value="2018">2018</option>
		<option value="2019">2019</option>
		<option value="2020">2020</option>
		<option value="2021">2021</option>
		<option value="2022">2022</option>
		<option value="2023">2023</option>
		<option value="2024">2024</option>
		<option value="2025">2025</option>
		<option value="2026">2026</option>
		<option value="2027">2027</option>
		<option value="2028">2028</option>
		</select>�� <select size="1" name="season_season">
		<option value="1">1</option>
		<option value="2">2</option>
		<option value="3">3</option>
		<option value="4">4</option>
		</select>���� <input type="submit" value="�ύ" name="B2"></td>
	</tr>
	</form>
	<!--�������ͳ�Ʊ�����  //star -->
	<form name="form5" method="post" action="order_info_SaleCountDetail.asp" id="form5">
    <input type="hidden" name="action" value="count_year"> 
	</form>
</tbody>
</table>

</body>

</html>
 
