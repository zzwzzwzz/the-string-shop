<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file="Sub.asp"-->
<!--#include file=include/Pages.asp-->
<%
dim search_prod_bid,search_prod_sid,search_prod_name,search_prod_detail,search_prod_UserPriceMin,search_prod_UserPriceMax,search_prod_flag
search_prod_bid         =my_request("bid",0)              '大类别id
search_prod_sid         =my_request("sid",0)              '小类别id
search_prod_name        =my_request("name",0)             '商品名称关键字
search_prod_detail		=my_request("detail",0)           '商品介绍关键字
search_prod_UserPriceMin=my_request("UserPriceMin",0)     '商品价格下限
search_prod_UserPriceMax=my_request("UserPriceMax",0)     '商品价格上限
search_prod_flag		=my_request("flag",0)             '商品特性

cx=request("cx")
if cx="" then cx=1
Select case cx
case 3
    SortBy=" order by product_info_name asc"
case 2
    SortBy=" order by product_info_PriceS asc"
case 1
    SortBy=" order by Addtime desc"
case else
    SortBy=" order by addtime desc"
end select

showlist=request("showlist")
if showlist="" then showlist=1

//参数设置表中相关参数调出
dim rs,sql,root_option_NumsPerRow,root_option_WidthSPic,root_option_HeighSPic,root_option_RowsPerPage,NumsPerPage
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_NumsPerRow,root_option_WidthSPic,root_option_HeighSPic,root_option_RowsPerPage from root_option where id=1"
rs.open sql,conn,1,1
root_option_NumsPerRow  =rs(0)
root_option_WidthSPic   =rs(1)
root_option_HeighSPic   =rs(2)
root_option_RowsPerPage =rs(3)
rs.close
set rs=nothing

if root_option_WidthSPic="" then root_option_WidthSPic=80
if root_option_HeighSPic="" then root_option_HeighSPic=80

NumsPerPage=root_option_NumsPerRow*root_option_RowsPerPage
if NumsPerPage="" then NumsPerPage=20
if NumsPerPage="0" then NumsPerPage=20
if root_option_NumsPerRow="" then root_option_NumsPerRow=5

dim Search
Search=""
if search_prod_bid<>"" then
    Search=Search & " and bid="&search_prod_bid
end if
if search_prod_sid<>"" then
    Search=Search & " and sid="&search_prod_sid
end if
if search_prod_name<>"" then
    Search=Search & " and product_info_name like '%"&search_prod_name&"%'"
end if
if search_prod_detail<>"" then
    Search=Search & " and product_info_detail like '%"&search_prod_detail&"%'"
end if
if search_prod_UserPriceMin<>"" then
    Search=Search & " and product_info_PriceS>="&search_prod_UserPriceMin
end if
if search_prod_UserPriceMax<>"" then
    Search=Search & " and product_info_PriceS<="&search_prod_UserPriceMax
end if
if search_prod_flag<>"" then
    Search=Search & " and instr(product_info_flag,"&search_prod_flag&")>0"
end if

call up("搜索结果","搜索结果","<a href=Product_Search.asp>商品搜索</a> &raquo; 搜索结果")
%>
<tr><td>
<!--显示方式及排序方式区  //star -->
		    <table border="0" width="98%" cellpadding="2" style="border-collapse: collapse">
             <tr>
				<td>
				  <form action="" name="taxis1" method="get">
				  <input type=hidden name=bid value=<%=search_prod_bid%>>
                  <input type=hidden name=sid value=<%=search_prod_sid%>>
                  <input type=hidden name=name value=<%=search_prod_name%>>
                  <input type=hidden name=detail value=<%=search_prod_detail%>>
                  <input type=hidden name=UserPriceMin value=<%=search_prod_UserPriceMin%>>
                  <input type=hidden name=UserPriceMax value=<%=search_prod_UserPriceMax%>>
                  <input type=hidden name=flag value=<%=search_prod_flag%>>
                  <input type=hidden name=cx value=<%=cx%>>
				   显示方式：<input name="showlist" type="radio" value="1" class="radio" onClick="document.taxis1.submit();" <%if showlist=1 then response.write "checked disabled"%>>图片
                      <input name="showlist" type="radio" value="2" class="radio" onClick="document.taxis1.submit();" <%if showlist=2 then response.write "checked disabled"%>>列表
                      <input name="showlist" type="radio" value="3" class="radio" onClick="document.taxis1.submit();" <%if showlist=3 then response.write "checked disabled"%>>纯文字</td>
				  </form>
				<form action="" name="taxis" method="get">
				<td align="right">
				  <input type=hidden name=bid value=<%=search_prod_bid%>>
                  <input type=hidden name=sid value=<%=search_prod_sid%>>
                  <input type=hidden name=name value=<%=search_prod_name%>>
                  <input type=hidden name=detail value=<%=search_prod_detail%>>
                  <input type=hidden name=UserPriceMin value=<%=search_prod_UserPriceMin%>>
                  <input type=hidden name=UserPriceMax value=<%=search_prod_UserPriceMax%>>
                  <input type=hidden name=flag value=<%=search_prod_flag%>>
                  <input type=hidden name=showlist value=<%=showlist%>>
                  排序方式：<input name="cx" type="radio" value="1" class="radio" onClick="document.taxis.submit();" <%if cx=1 then response.write "checked disabled"%>>上架时间
                      <input name="cx" type="radio" value="2" class="radio" onClick="document.taxis.submit();" <%if cx=2 then response.write "checked disabled"%>>价格
                      <input name="cx" type="radio" value="3" class="radio" onClick="document.taxis.submit();" <%if cx=3 then response.write "checked disabled"%>>商品名
                </td>
                </form>
			  </tr>
		    </table>
            <!--显示方式及排序方式区 //end-->
</td></tr>

<%response.write "<tr><td>"
response.write "<tr><td>"
call Product_ListSearch(Search,root_option_NumsPerRow,NumsPerPage)
response.write "</td></tr>"
call down()
%></center>