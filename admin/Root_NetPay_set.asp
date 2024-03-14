<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=0
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/base64.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select base_NetPay_AlipayOnOff,base_NetPay_AlipayEmail,base_NetPay_AlipaySafeCode,base_NetPay_ChinaBankOnOff,base_NetPay_ChinaBankUserName,base_NetPay_ChinaBankPrivateKey,base_NetPay_IpayOnOff,base_NetPay_IpayUserName,base_NetPay_IpayPrivateKey,base_NetPay_NpsOnOff,base_NetPay_NpsUserName,base_NetPay_NpsPrivateKey,base_NetPay_PayPalOnOff,base_NetPay_PayPalEmail,base_NetPay_AlipayPartnerID from root_NetPay where base_NetPay_id=1"
rs.open sql,conn,1,1
base_NetPay_AlipayOnOff        =rs(0)
base_NetPay_AlipayEmail        =rs(1)
base_NetPay_AlipaySafeCode     =rs(2)
base_NetPay_ChinaBankOnOff     =rs(3)
base_NetPay_ChinaBankUserName  =rs(4)
base_NetPay_ChinaBankPrivateKey=rs(5)
base_NetPay_IpayOnOff          =rs(6)
base_NetPay_IpayUserName       =rs(7)
base_NetPay_IpayPrivateKey     =rs(8)
base_NetPay_NpsOnOff           =rs(9)
base_NetPay_NpsUserName        =rs(10)
base_NetPay_NpsPrivateKey      =rs(11)
base_NetPay_PayPalOnOff        =rs(12)
base_NetPay_PayPalEmail        =rs(13)
base_NetPay_AlipayPartnerID    =rs(14)
rs.close
set rs=nothing

base_NetPay_AlipayEmail        =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipayEmail))),chr(13)&chr(10),"<br>")
base_NetPay_AlipaySafeCode     =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipaySafeCode))),chr(13)&chr(10),"<br>")
base_NetPay_AlipayPartnerID     =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipayPartnerID))),chr(13)&chr(10),"<br>")
base_NetPay_ChinaBankUserName  =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_ChinaBankUserName))),chr(13)&chr(10),"<br>")
base_NetPay_ChinaBankPrivateKey=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_ChinaBankPrivateKey))),chr(13)&chr(10),"<br>")
base_NetPay_IpayUserName       =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_IpayUserName))),chr(13)&chr(10),"<br>")
base_NetPay_IpayPrivateKey     =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_IpayPrivateKey))),chr(13)&chr(10),"<br>")
base_NetPay_NpsUserName        =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_NpsUserName))),chr(13)&chr(10),"<br>")
base_NetPay_NpsPrivateKey      =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_NpsPrivateKey))),chr(13)&chr(10),"<br>")
base_NetPay_PayPalEmail        =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_PayPalEmail))),chr(13)&chr(10),"<br>")

action=my_request("action",0)
if action="save" then
    call save()
end if

action2=my_request("action2",0)
if action2="save" then
    call save2()
end if

action3=my_request("action3",0)
if action3="save" then
    call save3()
end if

action4=my_request("action4",0)
if action4="save" then
    call save4()
end if

action5=my_request("action5",0)
if action5="save" then
    call save5()
end if

'/֧����-����
sub save()
    base_NetPay_AlipayOnOff   =my_request("base_NetPay_AlipayOnOff",1)
    base_NetPay_AlipayEmail   =my_request("base_NetPay_AlipayEmail",0)
    base_NetPay_AlipaySafeCode=my_request("base_NetPay_AlipaySafeCode",0)
    base_NetPay_AlipayPartnerID=my_request("base_NetPay_AlipayPartnerID",0)
                
    if base_NetPay_AlipayOnOff=0 then
        if base_NetPay_AlipayEmail="" or base_NetPay_AlipaySafeCode="" OR base_NetPay_AlipayPartnerID="" then
            response.redirect "error.htm"
            response.end
        end if
    end if
    Set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from root_NetPay where base_NetPay_id=1"
    rs.open sql,conn,1,3
    rs("base_NetPay_AlipayOnOff")   =base_NetPay_AlipayOnOff
    rs("base_NetPay_AlipayEmail")   =strAnsi2Unicode(Base64encode(strUnicode2Ansi(base_NetPay_AlipayEmail)))
    rs("base_NetPay_AlipaySafeCode")=strAnsi2Unicode(Base64encode(strUnicode2Ansi(base_NetPay_AlipaySafeCode)))     
    rs("base_NetPay_AlipayPartnerID")=strAnsi2Unicode(Base64encode(strUnicode2Ansi(base_NetPay_AlipayPartnerID)))     
    rs.update
    rs.close
    set rs=nothing

    call ok("���ѳɹ�����֧�������ã�","Root_NetPay_Set.asp")
end sub

'PayPal-����
sub save5()
    base_NetPay_PayPalOnOff=my_request("base_NetPay_PayPalOnOff",1)
    base_NetPay_PayPalEmail=my_request("base_NetPay_PayPalEmail",0)
                
    if base_NetPay_PayPalOnOff=0 then
        if base_NetPay_PayPalEmail="" then
            response.redirect "error.htm"
            response.end
        end if
    end if
    Set rs=Server.CreateObject("ADODB.Recordset")
    sql="select * from root_NetPay where base_NetPay_id=1"
    rs.open sql,conn,1,3
    rs("base_NetPay_PayPalOnOff")        =base_NetPay_PayPalOnOff
    rs("base_NetPay_PayPalEmail")       =strAnsi2Unicode(Base64encode(strUnicode2Ansi(base_NetPay_PayPalEmail))) 
    
    rs.update
    rs.close
    set rs=nothing

    call ok("���ѳɹ�����PayPal���ã�","root_NetPay_set.asp")
end sub

conn.close
set conn=nothing
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����-����֧��-����</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language="JavaScript" type="text/JavaScript">
function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="none";
  }
  else
  {
   linkimg.style.display="";
  }
}

function showlist2(dd)
{
  if(dd=="a")
  {
   linkimg2.style.display="none";
  }
  else
  {
   linkimg2.style.display="";
  }
}

function showlist3(dd)
{
  if(dd=="a")
  {
   linkimg3.style.display="none";
  }
  else
  {
   linkimg3.style.display="";
  }
}

function showlist4(dd)
{
  if(dd=="a")
  {
   linkimg4.style.display="none";
  }
  else
  {
   linkimg4.style.display="";
  }
}

function showlist5(dd)
{
  if(dd=="a")
  {
   linkimg5.style.display="none";
  }
  else
  {
   linkimg5.style.display="";
  }
}
</script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
  <tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">����֧��- ����</td>
	</tr>
	<tr>
		<td colspan="2">���ٵ�����&nbsp; <a href="#1">֧����</a>&nbsp;&nbsp;
		<font color="#CCCCCC">&nbsp;| </font>&nbsp;&nbsp;
		<a href="#5">PayPal</a></td>
	</tr>
	<!--֧����// -->
	<form name="form1" action="Root_NetPay_Set.asp" method="post">
    <input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="altbg1"><a name="1"></a>֧��������</td>
	</tr>
	<tr>
		<td colspan="2">
		<a target="_blank" href="https://www.alipay.com/">
		<img border="0" src="../images/netpaylogo/NetPay_logo_alipay.gif" width="270" height="49"></a></td>
	</tr>
	<tr>
		<td>֧�������ÿ��أ�</td>
		<td>
		    <input type="radio" value="0" name="base_NetPay_AlipayOnOff" <%if base_NetPay_AlipayOnOff=0 then response.write "checked" %> onClick='showlist("b");'>����&nbsp;&nbsp; 
		    <input type="radio" value="1" name="base_NetPay_AlipayOnOff" <%if base_NetPay_AlipayOnOff=1 then response.write "checked" %> onClick='showlist("a");'>�ر�			</td>
	</tr>
	<tr id="linkimg" <%if base_NetPay_AlipayOnOff=1 then%>style='display:none'<%end if%>>
		<td>֧�����˻����ã�</td>
		<td>�������䣺<input type="text" name="base_NetPay_AlipayEmail" size="40" value="<%=base_NetPay_AlipayEmail%>"><br>
		�� ȫ �룺<input type="text" name="base_NetPay_AlipaySafeCode" size="40" value="<%=base_NetPay_AlipaySafeCode%>"><br>
		����������ID��<input type="text" name="base_NetPay_AlipayPartnerID" size="40" value="<%=base_NetPay_AlipayPartnerID%>"></td>
	</tr>
	<tr>
		<td>��</td>
		<td><input type="submit" value="�ύ" name="B1">&nbsp;
		<input type="reset" value="����" name="B2"></td>
	</tr>
	</form>
	<!--PayPal// -->
	<form name="form5" action="Root_NetPay_Set.asp" method="post">
    <input type="hidden" name="action5" value="save">
	<tr>
		<td colspan="2" class="altbg1"><a name="5"></a>PayPal����</td>
	</tr>
	<tr>
		<td colspan="2"><b><a href="http://www.paypal.com.cn">
		<img border="0" src="../images/netpaylogo/NetPay_logo_paypal.gif" width="200" height="50"></a></b></td>
	</tr>
	<tr>
		<td>���ÿ��أ�</td>
		<td>
		    <input type="radio" value="0" name="base_NetPay_PayPalOnOff" <%if base_NetPay_PayPalOnOff=0 then response.write "checked" %> onClick='showlist5("b");'>����&nbsp;&nbsp; 
		    <input type="radio" value="1" name="base_NetPay_PayPalOnOff" <%if base_NetPay_PayPalOnOff=1 then response.write "checked" %> onClick='showlist5("a");'>�ر�</td>
	</tr>
	<tr id="linkimg5" <%if base_NetPay_PayPalOnOff=1 then%>style='display:none'<%end if%>>
		<td>�˻����ã�</td>
		<td>�������䣺<input type="text" name="base_NetPay_PaypalEmail" size="30" value="<%=base_NetPay_PaypalEmail%>"></td>
	</tr>
	<tr>
		<td>��</td>
		<td><input type="submit" value="�ύ" name="B9">&nbsp;
		<input type="reset" value="����" name="B10"></td>
	</tr>
	</form>
  </tbody>
</table>
<br><br>
</body>

</html>
 
