<%
'<!----News top10  ---->		
response.write  "<table width=100% cellspacing=0 cellpadding=4 class=MainTable>"&_
				"	<tbody class=table_td><tr>"&_
				"		<td class=MainHead><a href=News_List.asp class=U>最新文章</a></td>"&_
				"		<td class=MainHead align=right><a href=News_List.asp class=U><span style='font-weight: 400'>更多文章</span></a></td>"&_
				"	</tr>"
				
				'最新10条文章调出
					set rs=server.createobject("adodb.recordset")
					sql="select top 7 id,news_info_title,news_info_type,news_info_content from news_info order by id desc"
					rs.open sql,conn,1,1
					if not rs.eof then 
    					set id                =rs(0)
    					set news_info_title   =rs(1)
    					set news_info_type    =rs(2)
    				
        					set news_info_content =rs(3)
						
    			
    
    					while not rs.eof 
    					if len(news_info_title)>22 then 
        					news_info_title1=left(news_info_title,20)&"..."
    					else
        					news_info_title1=news_info_title
    					end if
    					response.write "<tr><td colspan=2>"
    
    					if news_info_type=1 then
        					response.write "·<a href=News_Detail.asp?id="&id&">"&news_info_title1&"</a>"
   					 	else
        					response.write "·<a href="&news_info_content&" target=_blank>"&news_info_title1&"</a>"
    					end if
    					response.write "</td></tr>"
    					rs.movenext
    					wend
					end if
					rs.close
					set rs=nothing
response.write "</tbody></table>"
%>