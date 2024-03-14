<%
dim url
url=request.ServerVariables("SCRIPT_NAME") 
if(len(trim(request.ServerVariables("QUERY_STRING")))>0) then 
  url=url & "?" & request.ServerVariables("QUERY_STRING") 
end if

'<!----product class  ---->
set rs=server.createobject("adodb.recordset")
sql="select root_option_NumsPerRowSclass from root_option where id=1"
rs.open sql,conn,1,1
root_option_NumsPerRowSclass=rs(0)
rs.close
set rs=nothing
if root_option_NumsPerRowSclass=2 then		
	response.write  "<table width='100%' cellspacing=1 cellpadding=2 class=MainTable></table>"&_
					"<div class=brclass></div>"
	response.write  "<table width='100%' cellspacing=0 cellpadding=4 class=category_table>"&_
					"<tr><td class=MainHead colspan=2>��Ʒ����</td></tr>"
					Set rs= Server.CreateObject("ADODB.Recordset")
					sql="select prod_BigClass_id,prod_BigClass_name from prod_BigClass order by prod_BigClass_sort asc"
					rs.open sql,conn,1,1
					if not rs.eof then
    					set prod_BigClass_id=rs(0)
    					set prod_BigClass_name=rs(1)
    					while not rs.eof
    						response.write "<tr><td colspan=2><img src=images/icon_arrow_blue.gif> <a href=Product_ListCategory.asp?Bid="&prod_BigClass_id&" class=left_bid><b>"&prod_BigClass_Name&"</b></a></td></tr>"
    						//����С���
    						set rs_s=server.CreateObject("adodb.recordset")
							sql_s="select prod_SmallClass_id,prod_SmallClass_name,prod_SmallClass_bid from prod_SmallClass where prod_SmallClass_Bid=" & prod_BigClass_id & " order by prod_SmallClass_id"
    						rs_s.open sql_s,conn,1,1
    						if not rs_s.eof then
        						set prod_SmallClass_id=rs_s(0)
        						set prod_SmallClass_name=rs_s(1)
        						set prod_SmallClass_bid=rs_s(2)
        						i=1
       							while not rs_s.eof
        						response.write "<td>&nbsp;&nbsp;<a href=Product_ListCategory.asp?Bid="&prod_SmallClass_Bid&"&Sid="&prod_SmallClass_id&">"&prod_SmallClass_name&"</a></td>"
	    						if (i mod 2)=0 then
	    							response.write "</tr>"
	  							end if
	  							rs_s.movenext
	  							i=i+1
	    						wend
							end if
							rs_s.close
							set rs_s=nothing
						rs.movenext
						wend
					end if
					rs.close
					set rs=nothing 
	response.write  "</table>"&_
			"<div class=brclass></div>"
else
	response.write  "<table width='100%' cellspacing=1 cellpadding=4 class=MainTable><tbody class=table_td>"&_
					"	<tr><td class=MainHead>��Ʒ����</td></tr>"
					Set rs= Server.CreateObject("ADODB.Recordset")
					sql="select prod_BigClass_id,prod_BigClass_name from prod_BigClass order by prod_BigClass_sort asc"
					rs.open sql,conn,1,1
					if not rs.eof then
    					set prod_BigClass_id=rs(0)
    					set prod_BigClass_name=rs(1)
    					while not rs.eof
    						response.write "<tr><td><img src=images/icon_arrow_blue.gif> <a href=Product_ListCategory.asp?Bid="&prod_BigClass_id&" class=left_bid><b>"&prod_BigClass_Name&"</b></a></td></tr>"
    						//����С���
    						set rs_s=server.CreateObject("adodb.recordset")
							sql_s="select prod_SmallClass_id,prod_SmallClass_name,prod_SmallClass_bid from prod_SmallClass where prod_SmallClass_Bid=" & prod_BigClass_id & " order by prod_SmallClass_id"
    						rs_s.open sql_s,conn,1,1
    						if not rs_s.eof then
        						set prod_SmallClass_id=rs_s(0)
        						set prod_SmallClass_name=rs_s(1)
        						set prod_SmallClass_bid=rs_s(2)
       							while not rs_s.eof
        						response.write "<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Product_ListCategory.asp?Bid="&prod_SmallClass_Bid&"&Sid="&prod_SmallClass_id&">"&prod_SmallClass_name&"</a></td></tr>"
	  							rs_s.movenext
	    						wend
							end if
							rs_s.close
							set rs_s=nothing
						rs.movenext
						wend
					end if
					rs.close
					set rs=nothing 
	response.write  "</tbody></table>"&_
					"<div class=brclass></div>"
end if

'<!----hot top10  ---->
response.write  "<table width=100% cellspacing=0 cellpadding=4 class=MainTable>"&_
				"	<tbody class=table_td><tr>"&_
				"		<td class=MainHead>��������</a></td>"&_
				"		<td class=MainHead align=right><a href=News_List.asp class=U><span style='font-weight: 200'>��������</span></a></td>"&_
				"	</tr>"
'����10�����µ���
set rs = server.createobject("ADODB.Recordset")
sql = "select top 10 id, news_info_title, news_info_content from news_info order by id desc"
rs.open sql, conn, 1, 1
if not rs.eof then
    dim news_info_title, news_info_content, news_info_addtime, news_info_hitnums
    set news_info_id      = rs(0)
    set news_info_title   = rs(1)
    set news_info_content = rs(2)
    while not rs.eof
        if len(news_info_title) > 22 then 
            news_info_title1 = left(news_info_title, 20)&"..."
        else
            news_info_title1 = news_info_title
        end if
        response.write "<tr><td colspan=2>"
        response.write "��<a href=News_Detail.asp?id="&news_info_id&">"&news_info_title1&"</a>"
        response.write "</td></tr>"
        rs.movenext
    wend
end if
rs.close
set rs=nothing
response.write "</tbody></table>"&_
				"<div class=brclass></div>"


'<!----  vote  ---->
set rs=server.createobject("adodb.recordset")
sql="select base_vote_OnOff from base_vote where base_vote_flag=1"
rs.open sql,conn,1,1
base_vote_OnOff=rs(0)
rs.close
set rs=nothing

if base_vote_OnOff=0 then
	response.write  "<table width=100% cellspacing=1 cellpadding=4 class=MainTable><tbody class=table_td>"&_
					"<tr><td class=MainHead>ͶƱ����</td></tr>"&_
					"<form action=votes.asp?vflag=add method=post target=win onSubmit=windowOpener()>"
   				 	sql="select base_vote_detail from base_vote where base_vote_flag=1"
    			 	set rs=conn.execute (sql)
    			 	base_vote_title=rs(0)
    			 	rs.close
    			 	set rs=nothing
	response.write  "<tr><td align=left>"&base_vote_title&"</td></tr>"&_
    				"<tr><td>"
    				sql="select base_vote_id,base_vote_detail from base_vote where base_vote_flag=0"
    				set rs=conn.execute (sql)
    				if not rs.eof then
        				set base_vote_id=rs(0)
        				set base_vote_detail=rs(1)
        				do while not rs.eof
	response.write  "	<input type=radio value="&base_vote_id&" name=idnums>"&base_vote_detail&"<br>"
        				rs.movenext
        				loop
    				end if
    				rs.close
    				set rs=nothing
	response.write  "</td></tr>"&_
    				"<tr><td align=center><input class=button type=submit value=ͶƱ���鿴���></td></tr>"&_
    				"</form></tbody></table><div class=brclass></div>"
end if
%>
