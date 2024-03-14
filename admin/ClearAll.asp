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
    call ok("数据库中所有数据已清空！","ClearAll.asp")
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>清空数据</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<h2 align="center"><br>
清空数据库中所有数据</h2>
<div align="center">
	<table border="0" width="50%" id="table1" cellpadding="4" style="border-collapse: collapse">
	<form name=form1 action=ClearAll.asp method=post>
	<input type="hidden" name="action" value="save"> 
	<tr>
			<td><font color="#FF0000">请谨慎操作,请确认你是否真的要清空数据库中所有数据信息,一般建议在首次进行网店设置时进行此项操作!</font></td>
		</tr>
		<tr>
			<td>
			<p align="center"> 
        <input type="submit" name="submit2" value="点此清空数据库中所有数据" onclick="{if(confirm('清空后将无法恢复，您确定要清空清空数据库中所有数据吗？')){this.document.form1.submit();return true;}return false;}" ></td>
		</tr>
	</form>
	</table>
</div>

</body>

</html>