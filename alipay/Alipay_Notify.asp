<!--#include file="Alipay_md5.asp"-->
<%
Dim key
'Partner 和 交易安全校验码
partner=""   'partner合作伙伴id
key =  ""  'partner账户的支付宝安全校验码
'ATN 校验地址 
'*******************************************************************
alipayNotifyURL = "https://www.alipay.com/cooperate/gateway.do?"
'获取ATN结果,如果你的服务器不支持https访问的话，需要用老的接口查询地址

alipayNotifyURL	= alipayNotifyURL & "service=notify_verify&partner=" & partner & "&notify_id=" & request.Form("notify_id")
	
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing


'*******************************************************************

'获取支付宝POST过来通知消息
For Each varItem in Request.Form 
mystr=varItem&"="&Request.Form(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
mystr=Left(mystr,Len(mystr)-1)
End If 
'response.write mystr
'*******************************************************************
mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'对参数排序
For i = Count TO 0 Step -1
minmax = mystr( 0 )
minmaxSlot = 0
For j = 1 To i
mark = (mystr( j ) > minmax)
If mark Then 
minmax = mystr( j )
minmaxSlot = j
End If 
Next
    
If minmaxSlot <> i Then 
temp = mystr( minmaxSlot )
mystr( minmaxSlot ) = mystr( i )
mystr( i ) = temp
End If
Next
'构造md5摘要字符串
 For j = 0 To Count Step 1
 value = SPLIT(mystr( j ), "=")

 If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  Then
 If j=Count Then
 md5str= md5str&mystr( j )
 Else 
 md5str= md5str&mystr( j )&"&"
 End If 
 End If 
 Next
md5str=md5str&key
 '生成md5摘要
 mysign=md5(md5str)

'*******************************************************************
	'验证消息的可靠性，并且处理自己的业务动作，然后反回给支付宝成功消息  
If mysign=request.Form("sign") And ResponseTxt="true" Then 	

  '判断支付状态，（文档中有支付枚举表，可供参考）
 If  request.Form("trade_status")="TRADE_FINISHED" Then 

'支付宝收到买家付款，请卖家发货 ,修改订单状态，发货等TRADE_FINISHED


response.write "success"
End If

Else
response.write "fail"
End If 

'*******************************************************************
 '写文本，方便测试（看网站需求，也可以改成存入数据库）
TOEXCELLR=TOEXCELLR&md5str&"MD5结果:"&mysign&"="&request.Form("sign")&"--ResponseTxt:"&ResponseTxt
set fs= createobject("scripting.filesystemobject") 
set ts=fs.createtextfile(server.MapPath("Notify_DATA/"&replace(now(),":","")&".txt"),true)

 ts.writeline(TOEXCELLR)
 ts.close
 set ts=Nothing
 set fs=Nothing




%>