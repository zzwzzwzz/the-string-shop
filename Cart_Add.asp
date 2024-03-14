<center><%
dim dbpath
dbpath=""
%>
<!--#include file="conn.asp"-->
<!--#include file="include/MyRequest.asp"-->
<%
id=my_request("id",1)
ProdNum=my_request("ProdNum",0)
if ProdNum="" then ProdNum=1

sql="select * from product_info where id="&id
set rs=conn.execute (sql)
if (rs.eof and rs.bof) then
    response.write "<script language=javascript>alert('您选择的商品不存在！');javascript:history.go(-1);</script>"
    response.End
end if
rs.close
set rs=nothing

'缺货否,缺货不能下单
dim Product_info_kucun
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select Product_info_kucun from Product_info where id="&id
rs.open sql,conn,1,1
Product_info_kucun  = rs(0)
rs.close
set rs=nothing
if product_info_kucun<=0 then
    response.write "<script language=javascript>alert('您选择的商品暂时缺货中,不能下单！');javascript:history.go(-1);</script>"
    response.End	
end if

if instr(session("ProdIds"),id)>0 then
    response.write "<script language=javascript>alert('您已经将此商品加入了购物车,请不要重复加入！');javascript:history.go(-1);</script>"
    response.End
end if

Sub PutToShopBag( add, x )
    If Len(x) = 0 Then
        x = add
    ElseIf InStr( x, add ) <= 0 Then
        x = x & "," & add
    End If
End Sub

ProdIds  = Session("ProdIds")
ProdNums = Session("ProdNums")

a = Split(id,",")
b = Split(ProdNum,",")

For I=0 To UBound(a)
    PutToShopBag a(I), ProdIds
    PutToShopBag b(I), ProdNums
Next

Session("ProdIds")  = ProdIds
Session("ProdNums") = ProdNums
Session.Timeout=30

//更新商品种数
y=session("y")
y=y+1
session("y")=y

response.redirect "Cart_List.asp"
%>

</center>