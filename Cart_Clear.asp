<%
Session.Contents.Remove("sum")
Session.Contents.Remove("ProdIds")
Session.Contents.Remove("ProdNums")
Session.Contents.Remove("ProdPrices")
Session.Contents.Remove("y")
response.redirect "Cart_List.asp"
response.End
%>


