<center><%dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp"-->
<!--#include file="include/base64.asp"-->
<!--#include file="include/mmd5.asp"-->
<%
'@ 接收页面示例

v_date   = Request.Form("v_date")   '接收订单日期
v_oid    = Request.Form("v_oid")    '接收订单编号
v_amount = Request.Form("v_amount") '接收订单金额
v_status = Request.Form("v_status") '接收订单状态
v_md5    = Request.Form("v_md5")    '接收MD5签名

//取得Ipay的私钥值
sql="select base_NetPay_IpayUserName,base_NetPay_IpayPrivateKey from base_NetPay where base_NetPay_id=1"
set rs=conn.execute (sql)
base_NetPay_IpayUserName  =rs(0)
base_NetPay_IpayPrivateKey=rs(1)
rs.close
set rs=nothing                              
v_mid=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_IpayUserName))),chr(13)&chr(10),"<br>")
v_key  =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_IpayPrivateKey))),chr(13)&chr(10),"<br>")

dim md5_string 
md5_string = asp_md5(cstr(v_date & v_mid & v_oid & v_amount & v_status & v_key)) 
 
'复核数字签名，采用MD5信息摘要算法 
v_md5 = md5(v_date & v_mid & v_oid & v_amount & v_status & v_key)
If v_md5 =  md5_string Then
	'支付结果可信	
	If  v_status = "00" Then
		Response.Write "实时支付成功！" 	'相应处理代码，核实支付金额是否相等，订单号是否存在，是否发过货等
		conn.execute("Update buyer set zt =1 where ddbh='"&v_oid&"'")
        conn.close
        set conn=nothing           
	ElseIf v_status =  "20" Then 
		Response.Write "历史支付成功！" 	'相应处理代码 
	Else 		 		
		Response.Write "支付失败！" 	'相应处理代码	
	End If	
Else  
	Response.Write "支付结果有误"		'支付结果有误
	'相应处理代码，设置COOKIE拒绝该客户使用，记录IP等 
End If	
%>

 
</center>