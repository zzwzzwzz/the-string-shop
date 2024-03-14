<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file="Sub.asp"-->
<!--#include file=include/Pages.asp-->
<%
Bid=Request("Bid")
Sid=Request("Sid")
if Sid="" then Sid=0

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

flag=request("flag")
if flag="" then flag=0
 
'������Ʒ��������
sql="select prod_BigClass_name from prod_BigClass where prod_BigClass_id="&Bid
set rs=conn.execute (sql)
BClass=rs(0)
rs.close
set rs=nothing

'������ƷС������
if Sid<>0 then
    sql="select prod_SmallClass_name from prod_SmallClass where prod_SmallClass_id="&Sid
    set rs=conn.execute (sql)
    SClass=rs(0)
    rs.close
    set rs=nothing
end if

if sid<>0 then
  txt_nav="<a href=Product_listCategory.asp?bid="&bid&"> "&Bclass&"</a> &raquo; "&SClass
  txt_title=SClass
else
  txt_nav=Bclass  
  txt_title=bclass
end if

'�������ñ�����ز�������
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_NumsPerRow,root_option_WidthSPic,root_option_HeighSPic,root_option_RowsPerPage from root_option where id=1"
rs.open sql,conn,1,1
root_option_NumsPerRow   =rs(0)
root_option_WidthSPic    =rs(4)
root_option_HeighSPic    =rs(5)
root_option_RowsPerPage  =rs(3)
rs.close
set rs=nothing

if root_option_WidthSPic="" then root_option_WidthSPic=130
if root_option_HeighSPic="" then root_option_HeighSPic=130

NumsPerPage=root_option_NumsPerRow*root_option_RowsPerPage
if NumsPerPage="" then NumsPerPage=20
if NumsPerPage="0" then NumsPerPage=20
if root_option_NumsPerRow="" then root_option_NumsPerRow=5

call up(txt_title&" �������Ʒ�б�",txt_title&" �������Ʒ�б�",txt_nav)

'��ʾ��ʽ������ʽ��
%>
<tr><td>  
			<!--��ʾ��ʽ������ʽ��  //star -->
		    <table border="0" width="100%" cellpadding="2" style="border-collapse: collapse">
             <tr>
				<td><form action="" name="taxis1" method="get">
				  <input type=hidden name=bid value=<%=bid%>>
                  <input type=hidden name=sid value=<%=sid%>>
                  <input type=hidden name=flag value=<%=flag%>>
				   ��ʾ��ʽ��<input name="showlist" type="radio" value="1" class="radio" onClick="document.taxis1.submit();" <%if showlist=1 then response.write "checked disabled"%>>ͼƬ
                      <input name="showlist" type="radio" value="2" class="radio" onClick="document.taxis1.submit();" <%if showlist=2 then response.write "checked disabled"%>>�б�
                      <input name="showlist" type="radio" value="3" class="radio" onClick="document.taxis1.submit();" <%if showlist=3 then response.write "checked disabled"%>>������</td>
				  </form>
				<td align="right"><form action="" name="taxis" method="get">
 				  <input type=hidden name=bid value=<%=bid%>>
                  <input type=hidden name=sid value=<%=sid%>>
                  <input type=hidden name=flag value=<%=flag%>>
                  <input type=hidden name=showlist value=<%=showlist%>>
                  ����ʽ��<input name="cx" type="radio" value="1" class="radio" onClick="document.taxis.submit();" <%if cx=1 then response.write "checked disabled"%>>�ϼ�ʱ��
                      <input name="cx" type="radio" value="2" class="radio" onClick="document.taxis.submit();" <%if cx=2 then response.write "checked disabled"%>>�۸�
                      <input name="cx" type="radio" value="3" class="radio" onClick="document.taxis.submit();" <%if cx=3 then response.write "checked disabled"%>>��Ʒ��
                </td>
                </form>
                <td align="right">
                <form action="" name="taxis2" method="get">
                <input type=hidden name=bid value=<%=bid%>>
                <input type=hidden name=sid value=<%=sid%>>
                <input type=hidden name=showlist value=<%=showlist%>>
                <input type=hidden name=cx value=<%=cx%>>
				<select name=flag size=1 onchange="document.taxis2.submit();">
                <option value="0" <%if flag=0 then response.write "selected"%>>����������</option>
                <option value="1" <%if flag=1 then response.write "selected"%>>��������Ʒ</option>
                <option value="2" <%if flag=2 then response.write "selected"%>>�������Ƽ�</option>
                <option value="3" <%if flag=3 then response.write "selected"%>>�������ؼ�</option>
                </select>
                </form></td>
			 </tr>
		    </table>
            <!--��ʾ��ʽ������ʽ�� //end-->
</td></tr>           
<tr><td>
<%call Product_ListCategory(bid,sid,root_option_NumsPerRow,NumsPerPage)%>
</td></tr>
<%call down()%></center>