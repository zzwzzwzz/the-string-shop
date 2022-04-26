<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=7
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>帮助信息-管理</title>
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
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组help_info_id中共有几维
        for v=1 to pp
            id=request("id")(v)
            conn.execute ("delete from [help_info] where id="&id)
        next
        call ok("所选信息已成功删除！","help_info_List.asp")
    end if
end sub
%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="3" class="header">帮助信息-管理</td>
	</tr>
    <tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<tr class="altbg1">
		<td>选中</td>
		<td>信息标题</td>
		<td>编辑</td>
	</tr>
	<form name="form1" action="help_info_List.asp" method="post">
    <%
    set rs=server.createobject("adodb.recordset")
    sql="select id,help_info_title from help_info order by id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=3 align=center>暂无相关帮助信息,<a href=help_info_Add.asp>请添加!</a></td></tr>"
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
      	
      	set id                = rs(0)
      	set help_info_title   = rs(1)
      	
      	while not rs.eof and i<=rs.pagesize
    %>
	<tr>
		<td><input type="checkbox" name="id" value="<%=id%>">   </td>
		<td><a href=Help_Info_Modi.asp?id=<%=id%>><%=help_info_title%></a></td>
		<td><a href=Help_Info_Modi.asp?id=<%=id%>>编辑</a></td>
	</tr>
	<%
        rs.movenext
        i=i+1
        wend
    %>
	<tr>
		<td colspan="3">
		<input type="checkbox" name="chkall" onclick="CheckAll(this.form)">全选 
        <input type="submit" name="action" value="删除" onclick="{if(confirm('删除后将无法恢复，您确定要删除选定的信息吗？')){this.document.form1.submit();return true;}return false;}">&nbsp;
        <input type="button" value="帮助信息-添加" name="action1" onclick="window.location='Help_Info_Add.asp'"></td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    <%
        call PageControl(iCount,maxpage,page)
    end if
    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
    %>
    </form>
</tbody>
</table>


</body>

</html>

