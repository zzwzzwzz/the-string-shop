<center>
<%dim dbpath
dbpath=""
%>
<!--#include file="conn.asp"-->
<!--#include file="include/MyRequest.asp"-->
<!--#include file="include/base64.asp"-->
<!--#include file="include/mmd5.asp"-->
<%
sql="select base_NetPay_NPSUserName,base_NetPay_NPSPrivateKey from base_NetPay where base_NetPay_id=1"
set rs=conn.execute (sql)
base_NetPay_NPSUserName  =rs(0)
base_NetPay_NPSPrivateKey=rs(1)
rs.close
set rs=nothing 
constPayEmail       =replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_NPSUserName))),chr(13)&chr(10),"<br>")
constPaySecurityCode=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_NPSPrivateKey))),chr(13)&chr(10),"<br>")

m_id			=	request("m_id")		
m_orderid		=	request("m_orderid")
m_oamount		=	request("m_oamount")
m_ocurrency		=	request("m_ocurrency")
m_url			=	request("m_url")	
m_language		=	request("m_language")
s_name			=	request("s_name")
s_addr			=	request("s_addr")
s_postcode		=	request("s_postcode")
s_tel			=	request("s_tel")
s_eml			=	request("s_eml")
r_name			=	request("r_name")
r_addr			=	request("r_addr")
r_postcode		=	request("r_postcode")
r_eml			=	request("r_eml")
r_tel			=	request("r_tel")
m_ocomment		=	request("m_ocomment")
m_status		=	request("m_status")
modate			=	request("modate")
newmd5info		=	request("newmd5info")
key				=	constPaySecurityCode

if request("md5info")="" then
    response.Write("认证签名为空值")
	response.end
end if

'升级的
newOrderMessage = m_id&m_orderid&m_oamount&key&m_status
newMd5text      = trim(md5(newOrderMessage))		

if Ucase(newMd5text)<>Ucase(newmd5info) then
	response.write("认证失败!!!")
else
	if m_status = 2 then
		response.write	("支付成功!")		&	"<br>"
        conn.execute("Update buyer set zt =1 where ddbh='"&m_orderid&"'")
        conn.close
        set conn=nothing
	else
		Response.Write "支付失败"
	end if
end if
%>

 

</center>