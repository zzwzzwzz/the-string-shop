<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=9
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言信息-管理</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
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
if action="删除" then
    call del()
end if

//过程：批量删除
sub del()
    guest_info_id=my_request("guest_info_id",0)
    if guest_info_id<>"" then
       pp=ubound(split(guest_info_id,","))+1 '判断数组id中共有几维
       for v=1 to pp
          id=request("guest_info_id")(v)     
          conn.execute ("delete from [guest_info] where guest_info_id="&id)
       next

      call ok("所选信息已成功删除！","GB_Info_List.asp")
    end if
end sub
%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">留言信息-管理</td>
	</tr>
	<tr>
		<td class="altbg1">选</td>
		<td class="altbg1">留言内容</td>
	</tr>
	<form name="form1" action="GB_Info_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select * from guest_info order by guest_info_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=2 align=center>目前暂无留言信息!</a></td></tr>"
    else
        rs.PageSize =10 '每页记录条数
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
        while not rs.eof and i<=rs.pagesize
    %>
	<tr>
		<td valign="top"><input type="checkbox" name="guest_info_id" value="<%=rs("guest_info_id")%>"></td>
		<td valign="top"><font color="#808080">用户名：</font><b><font color="#808080"><%=rs("guest_info_name")%></font></b><br>
		<font color="#808080">Email：</font><font color="#808080"><%=rs("guest_info_email")%></font><br>
		<font color="#808080">评论时间：</font><font color="#808080"><%=rs("guest_info_time")%></font><br>
		<font color="#808080">评论内容：</font><font color="#808080"><%=rs("guest_info_detail")%></font><hr>
		<%if rs("guest_info_backdetail")<>"" then%><font color="#cccccc"><b>已回复：</b></font><font color="#025793"><%=rs("guest_info_backDetail")%> </font>
		<font color="#999999">(回复时间：<%=rs("guest_info_BackTime")%>)</font><input type="button" value="编辑回复" name="action1" onclick="window.location='GB_info_back.asp?guest_info_id=<%=rs("guest_info_id")%>'"><%else%><input type="button" value="回复" name="action1" onclick="window.location='GB_info_back.asp?guest_info_id=<%=rs("guest_info_id")%>'"><%end if%>
        </td>
	</tr>
	<%   rs.movenext
         i=i+1
         wend
    %>
	<tr>
		<td colspan="2">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选 
        <input type="submit" name="action" value="删除" onclick="{if(confirm('删除后将无法恢复，您确定要删除选定的信息吗？')){this.document.form1.submit();return true;}return false;}"></td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    <%
        call PageControl(iCount,maxpage,page)
    end if
    rs.close
    set rs=nothing
    %>
</form>
</tbody>
</table>



</body>

</html>

