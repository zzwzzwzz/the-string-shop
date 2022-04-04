<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file="Sub.asp"-->
<!--#include file=include/Pages.asp-->
<%
flag=Request("flag")
select case flag
    case 1
        main_title="全部商品"
    case 2
        main_title="精品推荐" 
    case 3
        main_title="特价商品" 
end select

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
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_NumsPerRow,root_option_WidthSPic,root_option_HeighSPic,root_option_RowsPerPage from root_option where id=1"
rs.open sql,conn,1,1
root_option_NumsPerRow   =rs(0)
root_option_WidthSPic    =rs(1)
root_option_HeighSPic    =rs(2)
root_option_RowsPerPage  =rs(3)
rs.close
set rs=nothing

if root_option_WidthSPic="" then root_option_WidthSPic=80
if root_option_HeighSPic="" then root_option_HeighSPic=80

NumsPerPage=root_option_NumsPerRow*root_option_RowsPerPage
if NumsPerPage="" then NumsPerPage=20
if NumsPerPage="0" then NumsPerPage=20
if root_option_NumsPerRow="" then root_option_NumsPerRow=5

call up(main_title,main_title,main_title)
%>
<tr><td>
			<!--显示方式及排序方式区  //star -->
		    <table border="0" width="100%" cellpadding="2" style="border-collapse: collapse">
             <tr>
				<td>
				  <form action="" name="taxis1" method="get">
				  <input type=hidden name=flag value=<%=flag%>>
                  显示方式：<input name="showlist" type="radio" value="1" class="radio" onClick="document.taxis1.submit();" <%if showlist=1 then response.write "checked disabled"%>>图片
                      <input name="showlist" type="radio" value="2" class="radio" onClick="document.taxis1.submit();" <%if showlist=2 then response.write "checked disabled"%>>列表
                      <input name="showlist" type="radio" value="3" class="radio" onClick="document.taxis1.submit();" <%if showlist=3 then response.write "checked disabled"%>>纯文字</td>
				  </form>
				<form action="" name="taxis" method="get">
				<td align="right">
				<input type=hidden name=flag value=<%=flag%>>
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
call Product_ListFlag(flag,root_option_NumsPerRow,NumsPerPage)
response.write "</td></tr>"
call down()
%></center>