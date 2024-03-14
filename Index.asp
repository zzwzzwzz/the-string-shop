<center><%dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp" -->
<!--#include file="sub.asp" -->
<%
'�������ñ�����ز�������
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_NumsPerRow,root_option_RowsIndexNew,root_option_RowsIndexTj,root_option_RowsIndexSpec,root_option_WidthSPic,root_option_HeighSPic from root_option where id=1"
rs.open sql,conn,1,1
root_option_NumsPerRow   =rs(0)
root_option_RowsIndexNew =rs(1)
root_option_RowsIndexTj  =rs(2)
root_option_RowsIndexSpec=rs(3)
root_option_WidthSPic    =rs(4)
root_option_HeighSPic    =rs(5)
rs.close
set rs=nothing

if root_option_NumsPerRow="" then root_option_NumsPerRow=5
if root_option_WidthSPic="" then root_option_WidthSPic=80
if root_option_HeighSPic="" then root_option_HeighSPic=80
if root_option_RowsIndexNew="" then root_option_RowsIndexNew=2
if root_option_RowsIndexTJ="" then root_option_RowsIndexTJ=2
if root_option_RowsIndexSpec="" then root_option_RowsIndexSpec=2

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_info_IndexTitle,root_info_IndexKeyWords,root_info_IndexDescription from root_info where id=1"
rs.open sql,conn,1,1
root_info_IndexTitle      =rs(0)
root_info_IndexKeyWords   =rs(1)
root_info_IndexDescription=rs(2)
rs.close
set rs=nothing

response.write  "<html>"&_
				"<head>"&_
				"<meta http-equiv=Content-Language content=zh-cn>"&_
				"<title>"&root_info_IndexTitle&"</title>"&_
				"<meta name=keywords content="&root_info_IndexKeyWords&">"&_
				"<meta name=description content="&root_info_IndexDescription&">"&_
				"</head>"&_

				"<body>"
%>
<!--#include file="top.asp" -->
<%
response.write  "<table border=0 width=100% cellpadding=0 style='border-collapse: collapse'>"&_
				"	<tr>"&_
				"		<td width=190 valign=top>"
				
%>		
<!--#include file="Left.asp"-->
<%
response.write  "		</td>"&_
				"		<td width=10></td>"&_
				"		<td valign=top>"&_
				"			<table border=0 width=100% cellpadding=0 style='border-collapse: collapse'>"&_
				"				<tr>"&_
				"					<td width=58% valign=top id=ss>"&_
				"					</td>"&_
				"					<td width=2% ></td>"&_
				"					<td width=40% valign=top>"
response.write  "					</td>"&_
				"				</tr>"&_
				"			</table><div class=brclass></div>"
							call ProductIndexList(2,root_option_NumsPerRow,root_option_RowsIndexNew)
							call ProductIndexList(1,root_option_NumsPerRow,root_option_RowsIndexTj)
response.write  "		</td>"&_
                "	</tr>"&_
                "</table>"&_
                "<div class=brclass></div>"
%>
<!--#include file="End.asp"-->			
<%
response.write  "</body>"&_
				"</html>"
%></center>