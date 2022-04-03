<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=3
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员-会员邮件群发-列表</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language = "JavaScript">   
//全选操作    
function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }
</script>
<%
action=my_request("action",0)
if action="群发邮件" then
    call send()
end if

sub send()
    user_info_id=my_request("user_info_id",0)
    response.redirect ("User_email_send.asp?user_info_id="&user_info_id&"")
    response.end
end sub

%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="4" class="header">邮件群发-会员邮件列表</td>
	</tr>
	<tr>
		<td class="altbg1">选中</td>
		<td class="altbg1">会员用户名</td>
		<td class="altbg1">真实姓名</td>
		<td class="altbg1">会员Email</td>
	</tr>
	<form name="form1" action="user_email_send.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select user_info_id,user_info_UserName,user_info_RealName,user_info_email from user_info order by user_info_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=4 align=center>目前暂无会员邮件信息!</a></td></tr>"
    else
        rs.PageSize =20 '每页记录条数
        iCount=rs.RecordCount '记录总数
        iPageSize=rs.PageSize
        maxpage=rs.PageCount 
        page=request("page")  
        if Not IsNumeric(page) or page="" then
            page=1
        else
            page=cint(page)
        end if    
        if page<1 then
            page=1
        elseif  page>maxpage then
            page=maxpage
        end if   
        rs.AbsolutePage=Page
        if page=maxpage then
	        x=iCount-(maxpage-1)*iPageSize
        else
	        x=iPageSize
        end if
        i=1
        
        set user_info_id=rs(0)
        set user_info_UserName=rs(1)
        set user_info_RealName=rs(2)
        set user_info_email=rs(3)
        while not rs.eof and i<=rs.pagesize
    %>

	<tr>
		<td><input type="checkbox" name="user_info_id" value="<%=user_info_id%>"></td>
		<td><%=user_info_UserName%></td>
		<td><%=user_info_RealName%></td>
		<td><%=user_info_email%></td>
	</tr>
	<%
         rs.movenext
         i=i+1
         wend
    %>
	<tr>
		<td colspan="4">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选 
        <input type="submit" value="群发邮件" name="action"></td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    </form>
</tbody>
</table>
    <p align="center"><font color="#C0C0C0">
    <%
        call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
    end if
    rs.close
    set rs=nothing
    %>

</body>

</html>
 
