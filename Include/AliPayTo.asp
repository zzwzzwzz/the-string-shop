<%
	'������յĹ���url
	Public Function creatAlipayURL(t1,t4,t5,service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,seller_email,notify_url,key)
		Dim itemURL
		dim INTERFACE_URL,imgsrc,imgtitle
		'��ʼ������Ҫ����
		INTERFACE_URL	= t1	'֧���ӿ�
		imgsrc			= t4		'֧������ťͼƬ
		imgtitle		= t5		'��ť��ͣ˵��
  	
		mystr = Array("service="&service,"agent="&agent,"partner="&partner,"subject="&subject,"body="&body,"out_trade_no="&out_trade_no,"price="&price,"discount="&discount,"show_url="&show_url,"quantity="&quantity,"payment_type="&payment_type,"logistics_type="&logistics_type,"logistics_fee="&logistics_fee,"logistics_payment="&logistics_payment,"logistics_type_1="&logistics_type_1,"logistics_fee_1="&logistics_fee_1,"logistics_payment_1="&logistics_payment_1,"seller_email="&seller_email,"notify_url="&notify_url)
		Count=ubound(mystr)
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
	
		For j = 0 To Count Step 1
  			value = SPLIT(mystr( j ), "=")
  			If  value(1)<>"" then
       			If j=Count Then
       				md5str= md5str&mystr( j )
	   			Else 
       				md5str= md5str&mystr( j )&"&"
	   			End If 
  			End If 
  		Next

       md5str=md5str&key
	   sign=md5(md5str)
	   itemURL	= itemURL&INTERFACE_URL 

		For j = 0 To Count Step 1 
	    	value = SPLIT(mystr( j ), "=")
			If  value(1)<>"" then
				itemURL= itemURL&mystr( j )&"&"
			End If 	     
  		Next
		itemURL	= itemURL&"sign="&sign&"&sign_type="&sign_type   
		Response.Write   "<script language=""javascript"">"   
        Response.Write   "window.open('"&itemURL&"');"   
        Response.Write   "</script>"   
		'response.redirect itemURL             
	End Function
%>