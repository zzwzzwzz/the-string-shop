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
<title>留言-留言信息-回复</title>
<link rel="stylesheet" type="text/css" href="style.css">
<SCRIPT language="javascript" src="images/calendar.js"></SCRIPT> 
<SCRIPT language="javascript" src="images/datefunction.js"></SCRIPT>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">销售统计</td>
	</tr>
	<!--期间销售统计报表区  //star -->
    <tr>
		<td colspan="2" class="altbg1">期间销售统计：</td>
	</tr>
	<form name="form1" method="post" action="order_info_SaleCountDetail.asp" id="form1">
    <input type="hidden" name="action" value="count_time"> 
	<tr>
		<td>统计期间选择：</td>
		<td>从<input size="10"  name="Time_DayFrom" id="Time_DayFrom" onclick="event.cancelBubble=true;showCalendar('Time_DayFrom',false,'Time_DayFrom')" value="">&nbsp; 
		至 <input size="10"  name="Time_DayTo" id="Time_DayTo" onclick="event.cancelBubble=true;showCalendar('Time_DayTo',false,'Time_DayTo')" value="">
		<input type="submit" value="提交" name="B3"></td>
	</tr>
    </form>
    <!--日销售统计报表区  //star -->
    <form name="form2" method="post" action="order_info_SaleCountDetail.asp" id="form2">
    <input type="hidden" name="action" value="count_day"> 	
    <tr>
		<td colspan="2" class="altbg1">日销售统计：</td>
		</tr>
    <tr>
		<td>统计日期选择：</td>
		<td><input size="10"  name="day_day" id="day_day" onclick="event.cancelBubble=true;showCalendar('day_day',false,'day_day')" value=""> <input type="submit" value="提交" name="B1"></td>
	</tr>
	</form>
	<!--月销售统计报表区  //star -->
    <form name="form3" method="post" action="order_info_SaleCountDetail.asp" id="form2">
    <input type="hidden" name="action" value="count_month"> 	
    <tr>
		<td colspan="2" class="altbg1">月销售统计：</td>
	</tr>
	<tr>
		<td>统计月份选择：</td>
		<td><select size="1" name="month_year">
		<option value="2006" <%if year(now)=2006 then response.write "selected"%>>2006</option>
		<option value="2007" <%if year(now)=2007 then response.write "selected"%>>2007</option>
		<option value="2008" <%if year(now)=2008 then response.write "selected"%>>2008</option>
		<option value="2009" <%if year(now)=2009 then response.write "selected"%>>2009</option>
		<option value="2010" <%if year(now)=2010 then response.write "selected"%>>2010</option>
		<option value="2011" <%if year(now)=2011 then response.write "selected"%>>2011</option>
		<option value="2012" <%if year(now)=2012 then response.write "selected"%>>2012</option>
		<option value="2013" <%if year(now)=2013 then response.write "selected"%>>2013</option>
		<option value="2014" <%if year(now)=2014 then response.write "selected"%>>2014</option>
		<option value="2015" <%if year(now)=2015 then response.write "selected"%>>2015</option>
		<option value="2016" <%if year(now)=2016 then response.write "selected"%>>2016</option>
		</select>年 <select size="1" name="month_month">
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
		</select>月 <input type="submit" value="提交" name="B4"></td>
	</tr>
	</form>
	<!--季度销售统计报表区  //star -->
    <tr>
		<td colspan="2" class="altbg1">季度销售统计：</td>
	</tr>
	<form name="form4" method="post" action="order_info_SaleCountDetail.asp" id="form4">
    <input type="hidden" name="action" value="count_season"> 
    	<tr>
		<td>统计季度选择：</td>
		<td><select size="1" name="season_year">
		<option value="2006">2006</option>
		<option value="2007">2007</option>
		<option value="2008">2008</option>
		<option value="2009">2009</option>
		<option value="2010">2010</option>
		<option value="2011">2011</option>
		<option value="2012">2012</option>
		<option value="2013">2013</option>
		<option value="2014">2014</option>
		<option value="2015">2015</option>
		<option value="2016">2016</option>
		</select>年 <select size="1" name="season_season">
		<option value="1">1</option>
		<option value="2">2</option>
		<option value="3">3</option>
		<option value="4">4</option>
		</select>季度 <input type="submit" value="提交" name="B2"></td>
	</tr>
	</form>
	<!--年度销售统计报表区  //star -->
	<form name="form5" method="post" action="order_info_SaleCountDetail.asp" id="form5">
    <input type="hidden" name="action" value="count_year"> 
	</form>
</tbody>
</table>

</body>

</html>
 
