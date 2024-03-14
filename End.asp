<%
response.write 	"<div style=""width:100%;text-align:left;"">"&_
					"<table width=100% cellpadding=4 class=MainTable style='border-collapse: collapse'><tbody class=table_td>"&_
					"	<tr>"&_
					"		<td class=MainHead></td>"&_
					"		<td class=MainHead align=right></td>"&_
					"	</tr>"
response.write  "<div class=brclass></div>"&_
				"<table border=0 width=100% cellpadding=4 style='border-collapse: collapse' class='end_table'>"&_
				"	<tr>"&_
				"		<td align=center>"
						Set rs=Server.CreateObject("ADODB.Recordset")
						sql="select root_info_tel,root_info_email,root_info_sitename,root_info_address,root_info_zip from root_info where id=1"
						rs.open sql,conn,1,1
						root_info_tel      =rs(0)
						root_info_email    =rs(1)
						root_info_sitename =rs(2)
						root_info_address	=rs(3)
						root_info_zip		=rs(4)
						rs.close
						set rs=nothing
						if root_info_address<>""  then response.write "��ϵ��ַ��"&root_info_address
    					if root_info_zip<>""      then response.write "&nbsp; �ʱࣺ"&root_info_zip&"<br>"
    					if root_info_tel<>""      then response.write "��ϵ�绰��"&root_info_tel
    					if root_info_email<>""    then response.write "&nbsp; E-mail��"&root_info_email&"<br>"
response.write  "		Copyright &copy; "&year(now())&root_info_SiteName&" &nbsp; ��Ȩ����  "&_
    			"		</td>"&_
				"	</tr>"&_
				"	<tr></tr>"&_
				"</table>"&_
				"</div><br>"
%>	