function isDateString(sDate)
{	var iaMonthDays = [31,28,31,30,31,30,31,31,30,31,30,31]
	var iaDate = new Array(3)
	var year, month, day

	if (arguments.length != 1) return false
	iaDate = sDate.toString().split("-")
	if (iaDate.length != 3) return false
	if (iaDate[1].length > 2 || iaDate[2].length > 2) return false

	year = parseFloat(iaDate[0])
	month = parseFloat(iaDate[1])
	day=parseFloat(iaDate[2])

	if (year < 1900 || year > 2100) return false
	if (((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0)) iaMonthDays[1]=29;
	if (month < 1 || month > 12) return false
	if (day < 1 || day > iaMonthDays[month - 1]) return false
	return true
}

function stringToDate(sDate, bIgnore)
{	var bValidDate, year, month, day
	var iaDate = new Array(3)
	
	if (bIgnore) bValidDate = true
	else bValidDate = isDateString(sDate)
	
	if (bValidDate)
	{  iaDate = sDate.toString().split("-")
		year = parseFloat(iaDate[0])
		month = parseFloat(iaDate[1]) - 1
		day=parseFloat(iaDate[2])
		return (new Date(year,month,day))
	}
	else return (new Date(1900,1,1))
}
function checkAll(form)
{
	var returnvalue = true;
	var errorstr="";
	//����

	if(form.triptype[0].checked)
	{
		if(form.startCode1.value=="non" || form.destinationCode1.value=="non")
		{
			returnvalue=false;
			errorstr+="���������/�������\n";
		}
		
		if(form.startCode1.value== form.destinationCode1.value)
		{
			returnvalue=false;
			errorstr+="����/������в�����ͬ\n";
		}
		
		if(!isDateString(form.takeoffDate1.value))
		{
			returnvalue=false
			errorstr+="��������ȷ���ڸ�ʽ\n YYYY-MM-DD\n";
		}
		var d = new Date();
		var s = d.getFullYear()+"-"+(d.getMonth()+1)+"-"+d.getDate();
		if(stringToDate(form.takeoffDate1.value)<stringToDate(s))
		{
			returnvalue=false
			errorstr+="�������������Ч\n";
		}

	}
	//����
	else if(form.triptype[1].checked)
	{
		if(!isDateString(form.takeoffDate1.value) )
		{
			returnvalue=false;
			errorstr+="��������ȷ�ĳ�������\n YYYY-MM-DD\n";
		}
		if(!isDateString(form.takeoffDate2.value) )
		{
			returnvalue=false;
			errorstr+="��������ȷ�ķ�������\n YYYY-MM-DD\n";
		}
		if(stringToDate(form.takeoffDate1.value)<new Date())
		{
			returnvalue=false
			errorstr+="�������������Ч\n";
		}
		if(form.startCode1.value=="non" || form.destinationCode1.value=="non")
		{
			returnvalue=false;
			errorstr+="��������������\n";
		}
		if(form.startCode1.value== form.destinationCode1.value)
		{
			returnvalue=false;
			errorstr+="����/������в�����ͬ\n";
		}
		if(stringToDate(form.takeoffDate2.value)<stringToDate(form.takeoffDate1.value))
		{
			returnvalue=false
			errorstr+="����ʱ�䲻�����ڳ���ʱ��\n";
		}
	}
	//���
	else if(form.triptype[2].checked)
	{
		if(!isDateString(form.takeoffDate1.value) )
		{
			returnvalue=false;
			errorstr+="��������ȷ�ĵ�һ��������\n YYYY-MM-DD\n";
		}
		if(!isDateString(form.takeoffDate2.value) )
		{
			returnvalue=false;
			errorstr+="��������ȷ�ĵڶ���������\n YYYY-MM-DD\n";
		}
		if(!isDateString(form.takeoffDate3.value) )
		{
			returnvalue=false;
			errorstr+="��������ȷ�ĵ�����������\n YYYY-MM-DD\n";
		}
		if(stringToDate(form.takeoffDate1.value)<=new Date())
		{
			returnvalue=false
			errorstr+="�������������Ч\n";
		}
		if(form.startCode1.value=="non" || form.destinationCode1.value=="non" || form.startCode2.value=="non" || form.startCode3.value=="non" || form.destinationCode2.value=="non" || form.destinationCode3.value=="non")
		{
			returnvalue=false;
			errorstr+="��������ȷ�ĳ���/�������\n";
		}
		if(form.startCode1.value== form.destinationCode1.value || form.startCode2.value== form.destinationCode2.value || form.startCode3.value== form.destinationCode3.value)
		{
			returnvalue=false;
			errorstr+="����/������в�����ͬ\n";
		}
		if(stringToDate(form.takeoffDate2.value)<stringToDate(form.takeoffDate1.value))
		{
			returnvalue=false;
			errorstr+="�ڶ��������ڲ������ڵ�һ��������\n";
		}
		if(stringToDate(form.takeoffDate3.value)<stringToDate(form.takeoffDate2.value))
		{
			returnvalue=false
			errorstr+="�����������ڲ������ڵڶ���������\n";
		}
	}
	if (!returnvalue)
		alert(errorstr);

	c_hostname = location.hostname
	c_hostname = c_hostname.split(".")
	//if (c_hostname[0] != "www" && c_hostname[0] != "yoee" && c_hostname[0] != "yuee" && c_hostname[0] != ""){
	if (c_hostname[0] == "leiyu" || c_hostname[0] == "sinatravel" || c_hostname[0] == "hiwing" || c_hostname[0] == "china" || c_hostname[0] == "zhongsou" || c_hostname[0] == "24hotel"){
	  form.action= "http://" +c_hostname[0] +".yoee.net/waiting.jsp"
	}
	return returnvalue;
}

function checkAllNew(form)
{
	var returnvalue = true;
	var errorstr="";
	if(!isDateString(form.takeoffDate1.value) )
	{
		returnvalue=false;
		errorstr+="��������ȷ�ĳ�������\n YYYY-MM-DD\n";
	}
	if(form.takeoffDate2.value!="" ){
		form.triptype.value="2";
		if(!isDateString(form.takeoffDate2.value) )
		{
			returnvalue=false;
			errorstr+="��������ȷ�ķ�������\n YYYY-MM-DD\n";
		}
		else{
			if(stringToDate(form.takeoffDate2.value)<stringToDate(form.takeoffDate1.value))
			{
				returnvalue=false
				errorstr+="����ʱ�䲻�����ڳ���ʱ��\n";
			}
		}
	}
	else
		form.triptype.value="1";
	var d = new Date();
	var s = d.getFullYear()+"-"+(d.getMonth()+1)+"-"+d.getDate();
	if(stringToDate(form.takeoffDate1.value)<stringToDate(s))
	{
		returnvalue=false
		errorstr+="�������������Ч\n";
	}
	if(form.startCode1.value=="non" || form.destinationCode1.value=="non")
	{
		returnvalue=false;
		errorstr+="�����뵽�����\n";
	}
	if(form.startCode1.value== form.destinationCode1.value)
	{
		returnvalue=false;
		errorstr+="����/������в�����ͬ\n";
	}
	if (!returnvalue)
		alert(errorstr);

	c_hostname = location.hostname
	c_hostname = c_hostname.split(".")
	//if (c_hostname[0] != "www" && c_hostname[0] != "yoee" && c_hostname[0] != "yuee" && c_hostname[0] != ""){
	if (c_hostname[0] == "leiyu" || c_hostname[0] == "sinatravel" || c_hostname[0] == "hiwing" || c_hostname[0] == "china" || c_hostname[0] == "zhongsou" || c_hostname[0] == "24hotel"){
	  form.action= "http://" +c_hostname[0] +".yoee.net/waiting.jsp"
	}
	return returnvalue;
}

function openwin(page,size)
 {
   window.open(page,"newuser","toolbar=no,location=no,directories=no,status=no,scrollbars=yes,menubar=no,resizable=no,"+size);
 }