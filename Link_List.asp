<center><!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
call up("友情链接","友情链接","友情链接")

response.write  "<tr><td><b>文字链接</b></td></tr>"&_
				"<tr>"&_
				"	<td>"
						Set rs= Server.CreateObject("ADODB.Recordset")
						sql="select link_info_url,link_info_detail from link_info where link_info_type=0 order by id"
						rs.open sql,conn,1,1
						if not rs.eof then
							set link_info_url=rs(0)
   							set link_info_detail=rs(1)
							while not rs.eof
    						response.write "<a href="&link_info_url&" target=_blank>"&link_info_detail&"</a>&nbsp;"
        					rs.movenext
							wend
						else
							response.write "暂无文字链接!"
						end if
						rs.close
						set rs=nothing 
response.write  "	</td>"&_
				"</tr>"&_
				"<tr><td><b>图片链接</b></td></tr>"&_
				"<tr>"&_
				"	<td>"
						Set rs= Server.CreateObject("ADODB.Recordset")
						sql="select link_info_url,link_info_detail from link_info where link_info_type=1 order by id"
						rs.open sql,conn,1,1
						if not rs.eof then
   							i=1
							set link_info_url1=rs(0)
   							set link_info_detail1=rs(1)
							while not rs.eof 
							response.write "<a href="&link_info_url1&" target=_blank><img src=uploadpic/"&link_info_detail1&"></a>&nbsp;"
        					rs.movenext
							wend
						else
							response.write "暂无图片链接!"
						end if
						rs.close
						set rs=nothing 
response.write  "	</td>"&_
				"</tr>"

call down()
%></center>