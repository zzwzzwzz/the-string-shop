<center><%dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp" -->
<%
myAction=trim(request("myAction"))
if myAction="Del" then
    ID=trim(request("ID"))
   	call RemoveItem(clng(ID))
   
   	'更新商品种数
   	y=session("y")
   	y=y-1
   	session("y")=y
   
   	ProdIds  = Session("ProdIds")
   	ProdNums = Session("ProdNums") 
  
   	aaa=split(ProdNums,",")
   	bbb=split(ProdIds,",")

   	for i=0 to ubound(bbb)
       	set rs=server.createobject("adodb.recordset")
       	sql="select product_info_PriceS from product_info where id="&bbb(i)
       	rs.open sql,conn,1,1
       	if not rs.eof then
           	While Not rs.EOF
           	RMB=rs(0)

           	x=aaa(i)
           	if aaa(i)="" then x=1
          	sum1=sum1 + csng(rmb) * x
           	sum1=FormatNumber(sum1,2)
          	session("sum")=sum1

           	rs.MoveNext
           	Wend
        end if
    next 
   
    if y=0 then session("sum")=0
    response.write "<meta http-equiv=""refresh"" content=""0;url=cart_list.asp"">"
end if

Sub PutToShopBag( add, ProdIds )
  	If Len(ProdIds) = 0 Then
     	ProdIds = add
  	ElseIf InStr( ProdIds, add ) <= 0 Then
     	ProdIds = ProdIds & "," & add
  	End If
End Sub

SUB RemoveItem(ID)
  	dim i,intPos,ProdIdss,newSize
  	ProdIds = Session("ProdIds")
  	ProdIdss=split(ProdIds,",")
  	For i = 0 To UBound(ProdIdss)
    If clng(ProdIdss(i)) = ID Then
	  	intPos = i
	  	Exit For
    End If
  	Next
  	For i = intPos To UBound(ProdIdss) - 1
    If Not ProdIdss(i) = "" Then
	  	ProdIdss(i) = ProdIdss(i+1)
    End If
  	Next
  	newSize=UBound(ProdIdss)-1
  	redim preserve ProdIdss(newSize)
  	ProdIds=""
  	For q=0 To newsize
    	puttoshopbag ProdIdss(q),ProdIds
  	Next
  	Session("ProdIds") = ProdIds
  
  	dim s,ProdNumss
  	ProdNums = Session("ProdNums")
  	ProdNumss=split(ProdNums,",") 
  	For s = intPos To UBound(ProdNumss) - 1
    If Not ProdNumss(s) = "" Then
	  	ProdNumss(s) = ProdNumss(s+1)
    End If
  	Next
  	redim preserve ProdNumss(newSize)
  	ProdNums=""
  	For p=0 To newsize
    	puttoshopbag ProdNumss(p),ProdNums
  	Next
  	Session("ProdNums") = ProdNums 
  
End SUB
%>
</center>