<!--#include file="admin_check.asp"-->
<%dim dbpath
dbpath="../"
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    conn.execute ("update [root_info] set root_info_OffNote=null,root_info_LogoPic=null,root_info_tel=null,root_info_email=null,root_info_skin=null,root_info_indextitle=null,root_info_indexkeywords=null,root_info_indexdescription=null,root_info_aboutus=null,root_info_sitename=null,root_info_address=null,root_info_zip=null where id=1")
    conn.execute ("update [root_option] set root_option_NumsPerRow=0,root_option_RowsPerPage=0,root_option_RowsIndexNew=0,root_option_RowsIndexTj=0,root_option_RowsIndexSpec=0,root_option_WidthSPic=0,root_option_HeighSPic=0,root_option_OnOffAlipayButton=0,root_option_GuestOrderOnOff=0,root_option_MarkYuan=0 where id=1")
    conn.execute ("update [root_netpay] set base_NetPay_AlipayOnOff=0,base_NetPay_AlipayEmail=null,base_NetPay_AlipaySafeCode=null,base_NetPay_PayPalOnOff=0,base_NetPay_PayPalEmail=null where base_NetPay_id=1")
    conn.execute ("delete from [root_deliver]")
    conn.execute ("delete from [base_vote]")
    conn.execute ("delete from [prod_BigClass]")
    conn.execute ("delete from [prod_SmallClass]")
    conn.execute ("delete from [product_info]")
    conn.execute ("delete from [prod_review]")
    conn.execute ("delete from [prod_favorite]")
    conn.execute ("delete from [order_info]")
    conn.execute ("delete from [order_buy]")
    conn.execute ("delete from [user_info]")
    conn.execute ("delete from [news_info]")
    conn.execute ("delete from [help_info]")
    conn.execute ("delete from [guest_info]")
    call ok("���ݿ���������������գ�","ClearAll.asp")
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<h2 align="center"><br>
������ݿ�����������</h2>
<div align="center">
	<table border="0" width="50%" id="table1" cellpadding="4" style="border-collapse: collapse">
	<form name=form1 action=ClearAll.asp method=post>
	<input type="hidden" name="action" value="save"> 
	<tr>
			<td><font color="#FF0000">���������,��ȷ�����Ƿ����Ҫ������ݿ�������������Ϣ,һ�㽨�����״ν�����������ʱ���д������!</font></td>
		</tr>
		<tr>
			<td>
			<p align="center"> 
        <input type="submit" name="submit2" value="���������ݿ�����������" onclick="{if(confirm('��պ��޷��ָ�����ȷ��Ҫ���������ݿ�������������')){this.document.form1.submit();return true;}return false;}" ></td>
		</tr>
	</form>
	</table>
</div>

</body>

</html>