<!--#include file="Alipay_md5.asp"-->
<%
Dim key
'Partner �� ���װ�ȫУ����
partner=""   'partner�������id
key =  ""  'partner�˻���֧������ȫУ����
'ATN У���ַ 
'*******************************************************************
alipayNotifyURL = "https://www.alipay.com/cooperate/gateway.do?"
'��ȡATN���,�����ķ�������֧��https���ʵĻ�����Ҫ���ϵĽӿڲ�ѯ��ַ

alipayNotifyURL	= alipayNotifyURL & "service=notify_verify&partner=" & partner & "&notify_id=" & request.Form("notify_id")
	
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing


'*******************************************************************

'��ȡ֧����POST����֪ͨ��Ϣ
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
'�Բ�������
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
'����md5ժҪ�ַ���
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
 '����md5ժҪ
 mysign=md5(md5str)

'*******************************************************************
	'��֤��Ϣ�Ŀɿ��ԣ����Ҵ����Լ���ҵ������Ȼ�󷴻ظ�֧�����ɹ���Ϣ  
If mysign=request.Form("sign") And ResponseTxt="true" Then 	

  '�ж�֧��״̬�����ĵ�����֧��ö�ٱ����ɹ��ο���
 If  request.Form("trade_status")="TRADE_FINISHED" Then 

'֧�����յ���Ҹ�������ҷ��� ,�޸Ķ���״̬��������TRADE_FINISHED


response.write "success"
End If

Else
response.write "fail"
End If 

'*******************************************************************
 'д�ı���������ԣ�����վ����Ҳ���Ըĳɴ������ݿ⣩
TOEXCELLR=TOEXCELLR&md5str&"MD5���:"&mysign&"="&request.Form("sign")&"--ResponseTxt:"&ResponseTxt
set fs= createobject("scripting.filesystemobject") 
set ts=fs.createtextfile(server.MapPath("Notify_DATA/"&replace(now(),":","")&".txt"),true)

 ts.writeline(TOEXCELLR)
 ts.close
 set ts=Nothing
 set fs=Nothing




%>