<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=0
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
select case action 
    case ""
        call vote_list()
    
    case "vote_title_modi"
        call vote_title_modi()
        
    case "vote_OnOff"
        call vote_OnOff()
        
    case "vote_add"
        call vote_add()
        
    case "vote_addsave"
        call vote_addsave() 
               
    case "vote_modisave"
        call vote_modisave()
        
    case "vote_del"
        call vote_del()
end select

sub vote_title_modi()
    base_vote_TitleId=my_request("base_vote_TitleId",0)
    base_vote_title  =my_request("base_vote_title",0)
    sql="update base_vote set base_vote_detail='"&base_vote_title&"' where base_vote_flag=1 and base_vote_id="&base_vote_TitleId
    conn.execute (sql)
    call ok("您已成功修改了投票问题！","Root_vote_set.asp")
end sub

sub vote_addsave()
    base_vote_detail=my_request("base_vote_detail",0)
    sql="insert into base_vote (base_vote_detail) values ('"&base_vote_detail&"')"
    conn.execute (sql)
    call ok("您已成功添加了一条投票备选项信息！","Root_vote_set.asp")
end sub

sub vote_modisave()
    base_vote_id    =my_request("base_vote_id",1)
    base_vote_detail=my_request("base_vote_detail",0)
    sql="update base_vote set base_vote_detail='"&base_vote_detail&"' where base_vote_flag=0 and base_vote_id="&base_vote_id
    conn.execute (sql)
    call ok("您已成功修改了一条投票备选项信息！","Root_vote_set.asp")
end sub

sub vote_del()
    base_vote_id=my_request("base_vote_id",1)
    sql="delete from base_vote where base_vote_id="&base_vote_id
    conn.execute(sql)
    call ok("您已成功删除了一条投票备选项信息！","Root_vote_set.asp")
end sub

sub vote_OnOff()
    base_vote_TitleId=my_request("base_vote_TitleId",1)
    base_vote_OnOff  =my_request("base_vote_OnOff",1)
    sql="update base_vote set base_vote_OnOff='"&base_vote_OnOff&"' where base_vote_flag=1 and base_vote_id="&base_vote_TitleId
    conn.execute(sql)
    response.redirect "Root_vote_set.asp"
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>基本-投票调查-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<%
sub vote_list()
    sql="select base_vote_id,base_vote_detail,base_vote_OnOff from base_vote where base_vote_flag=1"
    set rs=conn.execute (sql)
    base_vote_titleid=rs("base_vote_id")
    base_vote_title  =rs("base_vote_detail")
    base_vote_OnOff  =rs("base_vote_OnOff")
    rs.close
    set rs=nothing
%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">投票调查-设置</td>
	</tr>
	<form action="Root_vote_set.asp" method=post>
	<input type="hidden" name="base_vote_TitleId" value="<%=base_vote_TitleId%>">
	<input type="hidden" name="action" value="vote_title_modi">
	<tr>
		<td>投票状态：</td>
		<td><font color=#FF3300><b><%if base_vote_OnOff=0 then response.write "<a href=Root_vote_set.asp?action=vote_OnOff&base_vote_OnOff=1&base_vote_titleid="&base_vote_titleid&">开启中</a>" else response.write "<a href=Root_vote_set.asp?action=vote_OnOff&base_vote_OnOff=0&base_vote_titleid="&base_vote_titleid&">已关闭</a>"%></b></font>  
		<font color="#808080"><span style="font-weight: 400">( </font>
		<font color="#999999"><font face="新宋体">←</font>点击切换开关状态</font><font color="#808080"> 
		)</font></span></td>
	</tr>
	<tr>
		<td>调查问题：</td>
		<td>
		<input type="text" name="base_vote_title" size="40" value="<%=base_vote_title%>"> 
		<input type="submit" value="修改" name="B3"></td>
	</tr>
	</form>
	<tr>
		<td>备 选 项：</td>
		<td>
		<table border="1" width="100%" id="table5" cellpadding="4" style="border-collapse: collapse" bordercolor="#CCCCCC">
			<tr>
				<td bgcolor="#654321"><b><font color="#FFFFFF">序号</font></b></td>
				<td bgcolor="#654321"><b><font color="#FFFFFF">备选项</font></b></td>
				<td bgcolor="#654321"><b><font color="#FFFFFF">修改保存</font></b></td>
				<td bgcolor="#654321"><b><font color="#FFFFFF">删除</font></b></td>
			</tr>
			<%
			sql="select base_vote_id,base_vote_detail from base_vote where base_vote_flag=0 order by base_vote_id desc"
			set rs=conn.execute (sql)
			if rs.eof then
				response.write "<tr><td colspan=4 align=center><font color=#FF0000>暂无投票备选项,请<a href=Root_vote_set.asp?action=vote_add>添加备选项</a></font></td></tr>"
			else
				i=1
				do while not rs.eof
				base_vote_id=rs("base_vote_id")
				base_vote_detail=rs("base_vote_detail")
			%>
			<form action=Root_vote_set.asp method=post>
			<input type="hidden" name="base_vote_id" value="<%=base_vote_id%>">
			<input type="hidden" name="action" value="vote_modisave">
            <tr>
				<td><input type="text" name="<%=i%>" size="2" value="<%=i%>"></td>
				<td>
				<input type="text" name="base_vote_detail" size="30" value="<%=base_vote_detail%>"></td>
				<td> <input type="submit" value="修改" name="B4"></td>
				<td> 
				<input type="button" value="删除" onclick="window.location='Root_vote_set.asp?base_vote_id=<%=base_vote_id%>&action=vote_del'" name="B5"></td>
			</tr>
			</form>
			<%
			    rs.movenext
			    i=i+1
			    loop
			end if
			rs.close
			set rs=nothing
			%>
		</table><br>
		<a href=Root_vote_set.asp?action=vote_add>添加备选项</a>
		</td>
	</tr>
</tbody>
</table>
<%end sub%>

<%sub vote_add()%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Root_vote_set.asp" method=post>
<input type="hidden" name="action" value="vote_addsave">
    <tr>
		<td colspan="2" bgcolor="#654321" class="header">备选项-添加</td>
	</tr>
	<tr>
		<td>备选项：</td>
		<td><input type="text" name="base_vote_detail" size="30"></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="提交" name="B6">&nbsp;<input type="reset" value="重置" name="B7"></td>
	</tr>
</form>
</tbody>
</table>
<%end sub%>

</body>

</html>
 
