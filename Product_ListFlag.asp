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
        main_title="ȫ����Ʒ"
    case 2
        main_title="��Ʒ�Ƽ�" 
    case 3
        main_title="�ؼ���Ʒ" 
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

//�������ñ�����ز�������
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
			<!--��ʾ��ʽ������ʽ��  //star -->
		    <table border="0" width="100%" cellpadding="2" style="border-collapse: collapse">
             <tr>
				<td>
				  <form action="" name="taxis1" method="get">
				  <input type=hidden name=flag value=<%=flag%>>
                  ��ʾ��ʽ��<input name="showlist" type="radio" value="1" class="radio" onClick="document.taxis1.submit();" <%if showlist=1 then response.write "checked disabled"%>>ͼƬ
                      <input name="showlist" type="radio" value="2" class="radio" onClick="document.taxis1.submit();" <%if showlist=2 then response.write "checked disabled"%>>�б�
                      <input name="showlist" type="radio" value="3" class="radio" onClick="document.taxis1.submit();" <%if showlist=3 then response.write "checked disabled"%>>������</td>
				  </form>
				<form action="" name="taxis" method="get">
				<td align="right">
				<input type=hidden name=flag value=<%=flag%>>
                <input type=hidden name=showlist value=<%=showlist%>>
                   ����ʽ��<input name="cx" type="radio" value="1" class="radio" onClick="document.taxis.submit();" <%if cx=1 then response.write "checked disabled"%>>�ϼ�ʱ��
                      <input name="cx" type="radio" value="2" class="radio" onClick="document.taxis.submit();" <%if cx=2 then response.write "checked disabled"%>>�۸�
                      <input name="cx" type="radio" value="3" class="radio" onClick="document.taxis.submit();" <%if cx=3 then response.write "checked disabled"%>>��Ʒ��
                </td>
                </form>
			 </tr>
		    </table>
            <!--��ʾ��ʽ������ʽ�� //end-->
</td></tr>

<%response.write "<tr><td>"
call Product_ListFlag(flag,root_option_NumsPerRow,NumsPerPage)
response.write "</td></tr>"
call down()
%></center>