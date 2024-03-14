<%
dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>投票结果</title>
</head>

<body>
<table border="1" width="100%" cellspacing="0" style="border-style:solid; border-width:0; padding:0px; border-collapse: collapse" bordercolor="#ACA793" cellpadding="0">
	<tr>
		<td>
		<table border="0" width="100%" cellpadding="4" style="border-collapse: collapse" cellspacing="1">
			<tr>
				<td bgcolor="#99CCFF">
				<p align="center">
				<span style="font-size: 12px; font-weight:700">投票结果</span></td>
			</tr>
<%
vflag=my_request("vflag",0)
if vflag="add" then
    idnums=my_request("idnums",1)
    if Request.ServerVariables("REMOTE_ADDR")=request.cookies("IPAddress") then
        response.write"<SCRIPT language=JavaScript>alert('感谢您的支持，您已经投过票了，请勿重复投票，谢谢！');"
        response.write"javascript:window.close();</SCRIPT>"
    end if

	if idnums="" or isnull(idnums) then
		response.write"<SCRIPT language=JavaScript>alert('对不起，请选择投票类别！');"
		response.write"javascript:window.close();</SCRIPT>"
	else	
        sql="update base_vote set base_vote_nums=base_vote_nums+1 where base_vote_id="&idnums
        conn.execute (sql)
        response.cookies("IPAddress")=Request.ServerVariables("REMOTE_ADDR")
    end if
%>
			<tr>
				<td><span style="font-size: 12px">感谢您的参与</span></td>
			</tr>
<%
end if

Set rs= Server.CreateObject("ADODB.Recordset")
sql2="select * from base_vote where base_vote_flag=0"
rs.open sql2,conn,1,1
'求出数
Amount=0
if rs.eof then
    Amount=0
else
    For I = 1 To rs.RecordCount
    Amount = Amount + rs("base_vote_nums")
    rs.MoveNext
    Next

    '将数据移到第一条
    rs.MoveFirst
%>
			<tr>
				<td><span style="font-size: 9pt">到目前为止结果如下(一共投票数<%=Amount%>)</span></td>
			</tr>
			<tr>
				<td><hr color="#ACA793" size="1"></td>
			</tr>
			<%
			do while not rs.eof
			Percent = Round((rs("base_vote_nums")/Amount)*100,2)
			%>
			<tr>
				<td><span style="font-size: 9pt"><%=rs("base_vote_detail")%>:<img src="images/poll.gif" width="<%= Percent * 3 %>" height="10">&nbsp;&nbsp;<%=Percent%>%&nbsp;&nbsp;得票数:<%=rs("nums")%></span></td>
			</tr>
			<%rs.movenext
			loop
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			%>
		</table>
		</td>
	</tr>
</table>
<%end if%>
<center><a href="javascript:self.close();"><font color="#000000" style="font-size: 12px">关闭窗口</font></a></center>
</body>

</html>
 
