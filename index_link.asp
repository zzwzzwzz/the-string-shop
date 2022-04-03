<%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select link_info_url,link_info_detail from link_info where link_info_IndexShow=0 order by id"
rs.open sql,conn,1,1
if not rs.eof then
	response.write 	"<div style=""width:100%;text-align:left;"">"&_
					"<table width=100% cellpadding=4 class=MainTable style='border-collapse: collapse'><tbody class=table_td>"&_
					"	<tr>"&_
					"		<td class=MainHead>友情链接</td>"&_
					"		<td class=MainHead align=right><a href=Link_List.asp ><span style='font-weight: 400'>更多友情链接</span></a></td>"&_
					"	</tr>"
	Set rs1= Server.CreateObject("ADODB.Recordset")
	sql1="select link_info_url,link_info_detail from link_info where link_info_type=1 and link_info_IndexShow=0 order by id"
	rs1.open sql1,conn,1,1
	if not rs1.eof then
		response.write "<tr><td colspan=2>"
		set link_info_url=rs1(0)
   		set link_info_detail=rs1(1)
		while not rs1.eof
    	response.write "<a href="&link_info_url&" target=_blank><img src=uploadpic/"&link_info_detail&"></a>&nbsp;"
        rs1.movenext
		wend
		response.write "</td></tr>"
	end if
	rs1.close
	set rs1=nothing 

	dim link_info_url,link_info_detail
	Set rs1= Server.CreateObject("ADODB.Recordset")
	sql1="select link_info_url,link_info_detail from link_info where link_info_type=0 and link_info_IndexShow=0 order by id"
	rs1.open sql1,conn,1,1
	if not rs1.eof then
		response.write "<tr><td colspan=2>"
		set link_info_url    =rs1(0)
   		set link_info_detail =rs1(1)
		while not rs1.eof
    	response.write "<a href="&link_info_url&" target=_blank>"&link_info_detail&"</a>&nbsp;"
        rs1.movenext
		wend
		response.write "</td></tr>"
	end if
	rs1.close
	set rs1=nothing
	response.write "</tbody></table></div>"
end if
rs.close
set rs=nothing

%>
