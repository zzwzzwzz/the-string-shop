<center><%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Response.Expires = -100
dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp" -->
<!--#include file="include/nosql.asp" -->
<!--#include file="Sub.asp" -->
<%
url=request.servervariables("http_referer")

ProdIds  = Session("ProdIds")
ProdNums = Session("ProdNums")

Sub PutToShopBag( mc, ProdIds )
    If Len(ProdIds) = 0 Then
        ProdIds =mc
    ElseIf InStr( ProdIds, mc ) <= 0 Then
        ProdIds = ProdIds & ","& mc
    End If
End Sub

'保存商品数量
If Request("cmdShow") = "Yes" Then
    ProdIds = ""
    a = Split(nosql(request("mc")), ",")
    For I=0 To UBound(a)
        if  a(I)="" then a(I)=1
        PutToShopBag a(I), ProdIds
    Next
    Session("ProdIds") = ProdIds
   
    ProdNums = ""
    b = Split(nosql(request("pbuynums")), ",")
    For I=0 To UBound(b)
        if b(I)="" then b(I)=1
        PutToShopBag b(I), ProdNums
    Next
    Session("ProdNums") = ProdNums
    Response.write "<meta http-equiv=""refresh"" content=""0;url=cart_list.asp"">"
End If

call up("购物车状态","购物车状态","购物车状态")
response.write  "<tr>"&_
				"		<table border=1 width=100% cellpadding=4 style='border-collapse: collapse' bordercolor=#DFDFDF>"&_
				"			<tr><td>商品名称</td><td>市场价格</td><td>本站价格</td><td>订购数量</td><td>小计</td><td>删除</td></tr>"&_
				"			<form action=Cart_List.asp method=post name=form1 onsubmit=return CheckFrom();>"&_
				"			<input type=hidden name=cmdShow value=Yes>"
    						if ProdIds<>"" then
      		    				aaa=split(ProdNums,",")
     		    				bbb=split(ProdIds,",")
    		    				Quatitys=split(Request("pbuynums"),",")
                				session("y")=ubound(bbb)+1
        	    				for i=0 to ubound(bbb)
        	    				set rs=server.createobject("adodb.recordset")
        	    				sql="select id,product_info_PriceM,product_info_PriceS,product_info_name from product_info where id="&bbb(i)
        	    				rs.open sql,conn,1,1
        	    				if rs.eof or rs.bof then
         	        				response.write "<tr><td colspan=6 align=center><a href='javascript:history.go(-1)'>&lt;&lt; 购物车为空，请返回选购商品</a></td></tr>"
        	    				else
        	        				set id=rs(0)
        	        				set product_info_PriceM=rs(1)
									set RMB=rs(2)
        	        				set product_info_name=rs(3)
        	        				While Not rs.EOF
        	        				Quatity = Quatitys(i)
        	        				If Quatity <>"" Then 
        	            				x=Quatity
        	        				else
        	            				x=aaa(i)
       	                				if aaa(i)="" then x=1
       	            				end if
       	            				sum1=sum1 + csng(rmb) * x
       	            				sum=FormatNumber(sum1,2,-1)
       	            				session("sum")=sum
response.write  "			<tr>"&_
				"				<td><input type=hidden name=mc value="&id&"><a href=Product_Detail.asp?id="&id&" target=_blank>"&product_info_name&"</a></td>"&_
				"				<td>￥"&FormatNumber(product_info_PriceM,2,-1)&"</td>"&_
				"				<td><font color=#FF0000>￥"&FormatNumber(Rmb,2,-1)&"</font></td>"&_
				"				<td><input name=pbuynums value="&x&" size=5 maxlength=5></td>"&_
				"				<td>￥"&FormatNumber((csng(rmb)*x),2,-1)&"</td>"&_
				"				<td><a href=Cart_Del.asp?MyAction=Del&id="&id&">删除</a></td>"&_
				"			</tr>"
            	    				rs.MoveNext
            	    				Wend
        	   					end if
        	  					rs.close
        	   					set rs=nothing
        	   					next
response.write  "			<tr><td colspan=6 align=right>合计金额：<span style='color:#FF6633;font-size:18px;'>￥"&sum&"</span></td></tr>"&_
				"			<tr>"&_
				"				<td colspan=6 align=center>"&_
				"    				<input class=button name=order type=submit onFocus=this.blur() value=修改数量>&nbsp;"&_
				"    				<input class=button type=button value=清空商品 onclick=window.location='cart_clear.asp'>&nbsp;"&_
				"    				<input class=button name=Submit type=button onclick=window.location='index.asp' value=返回主页 onFocus=this.blur()>&nbsp;"&_
				"    				<input class=button type=button value=结账付款 onclick=window.location='Cart_GuestOrderChk.asp'>"&_
				"    			</td>"&_
				"			</tr>"
    	    				else
        	    				response.write "<tr><td colspan=6 align=center><a href=index.asp>购物车为空，请返回选购商品!</a></td></tr>"
    	    				end if
response.write  "		</table>"&_
				"	</td>"&_
				"</tr>"
call down()
%></center>