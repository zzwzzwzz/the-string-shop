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
						sql="select root_info_ICP,root_info_tel,root_info_email,root_info_QQ,root_info_MSN,root_info_WangWang,root_info_sitename,root_info_address,root_info_zip,root_info_fax from root_info where id=1"
						rs.open sql,conn,1,1
						root_info_ICP1      =rs(0)
						root_info_tel1      =rs(1)
						root_info_email1    =rs(2)
						root_info_QQ1       =rs(3)
						root_info_MSN1      =rs(4)
						root_info_WangWang1 =rs(5)
						root_info_sitename1 =rs(6)
						root_info_address1	=rs(7)
						root_info_zip1		=rs(8)
						root_info_fax1		=rs(9)
						rs.close
						set rs=nothing

						if root_info_address1<>""  then response.write "联系地址："&root_info_address1
    					if root_info_zip1<>""      then response.write "&nbsp; 邮编："&root_info_zip1&"<br>"
    					if root_info_tel1<>""      then response.write "联系电话："&root_info_tel1
    					if root_info_email1<>""    then response.write "&nbsp; E-mail："&root_info_email1&"<br>"
    					if root_info_qq1<>"" 	   then response.write "QQ："&root_info_QQ1
    					if root_info_WangWang1<>"" then response.write "&nbsp; 淘宝旺旺："&root_info_WangWang1  
    					if root_info_msn1<>"" 	   then response.write "&nbsp; MSN："&root_info_msn1&"<br>"
response.write  "		Copyright &copy; "&year(now())&root_info_SiteName1&" 版权所有  <a href=http://www.miibeian.gov.cn/ target=_blank>"&root_info_ICP1&"</a>"&_
    			"		</td>"&_
				"	</tr>"&_
				"	<tr><td height=5></td></tr>"&_
				"</table>"&_
				"</div><br>"
%>	