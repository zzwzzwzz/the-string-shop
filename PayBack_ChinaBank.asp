<center><%dim dbpath
dbpath=""
%>
<!--#include file="conn.asp"-->
<!--#include file="include/MyRequest.asp"-->
<!--#include file="include/base64.asp"-->
<!--#include file="include/mmd5.asp"-->
<%
//提取表单参数
v_oid=request("v_oid")              '商户发送的v_oid定单编号 
v_pmode=request("v_pmode")	        '支付方式（字符串）     
v_pstatus=request("v_pstatus")      '支付状态：20（支付成功）| 30（支付失败）
v_pstring=request("v_pstring")      '支付结果：支付完成（当v_pstatus=20时）|失败原因（当v_pstatus=30时）
v_amount=request("v_amount")        '订单实际支付金额
v_moneytype=request("v_moneytype")  '订单实际支付币种
remark1=request("remark1")          '备注字段1
remark2=request("remark2")          '备注字段2
v_md5str=request("v_md5str")        ' Md5校验串

sql="select base_NetPay_ChinaBankUserName,base_NetPay_ChinaBankPrivateKey from base_NetPay where base_NetPay_id=1"
set rs=conn.execute (sql)
base_NetPay_ChinaBankUserName  =rs(0)
base_NetPay_ChinaBankPrivateKey=rs(1)
rs.close
set rs=nothing 
constPayEmail=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_ChinaBankUserName))),chr(13)&chr(10),"<br>")
key          =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_ChinaBankPrivateKey))),chr(13)&chr(10),"<br>")

if request("v_md5str")="" then
	response.Write("v_md5str：空值")
	response.end
end if

'md5校验
text = v_oid&v_pstatus&v_amount&v_moneytype&key
md5text = Ucase(trim(md5(text)))

'按md5检验情况输出结果 Ucase转换为大写
if md5text<>v_md5str then
    response.write("MD5 error")
else
    '逻辑处理
    if v_pstatus=20 then
	    response.write "支付成功"
        conn.execute("Update buyer set zt =1 where ddbh='"&v_oid&"'")
        conn.close
        set conn=nothing
    else
	    response.write "支付失败"
    end if
end if
'------------------------------------------------------------------------------
'提示：仅是对校验码校验通过不表示该支付结果是成功只意味着该信息是由网银传回
'校验成功需对传回的v_pstatus参数做判断，其中20都意味着支付成功，30表示支付失败
'如果商户涉及实时售卡，请对返回的金额与数据库中原始金额做大小判断，以防恶意行为
'------------------------------------------------------------------------------
%>

 
</center>