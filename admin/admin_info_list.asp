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
<title>管理员-管理人员信息-管理</title>
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
    admin_info_id=my_request("admin_info_id",0)
    if admin_info_id<>"" then
       pp=ubound(split(admin_info_id,","))+1 '判断数组id中共有几维
       for v=1 to pp
          id=request("admin_info_id")(v)     
          conn.execute ("delete from [admin_info] where admin_info_id="&id)
       next

      call ok("所选信息已成功删除！","admin_info_list.asp")
    end if
end sub

%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="15" class="header">管理人员信息-管理</td>
	</tr>	
    <tr>
		<td class="altbg2" colspan="15"></td>
	</tr>
	<tr>
		<td class="altbg1">选中</td>
		<td class="altbg1">真实姓名</td>
		<td class="altbg1">用户名</td>
		<td class="altbg1" colspan="10">管理权限分配情况</td>
		<td class="altbg1" align="center">修改</td>
		<td class="altbg1" align="center">密码修改</td>
	</tr>
	<form name="form1" action="admin_info_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select admin_info_id,admin_info_flag,admin_info_RealName,admin_info_UserName from admin_info order by admin_info_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=15 align=center>目前暂无管理人员信息,<a href=admin_info_add.asp>请添加!</a></td></tr>"
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
        
        dim admin_info_id,admin_info_flag,admin_info_RealName,admin_info_UserName
        set admin_info_id      =rs(0)
        set admin_info_flag    =rs(1)
        set admin_info_RealName=rs(2)
        set admin_info_UserName=rs(3)
        while not rs.eof and i<=rs.pagesize
    %>
	<tr>
		<td rowspan="2"><input type="checkbox" name="admin_info_id" value="<%=admin_info_id%>"></td>
		<td rowspan="2"><%=admin_info_RealName%></td>
		<td rowspan="2"><%=admin_info_UserName%></td>
		<td style="background-color: #F3F3F3">基本设置</td>
		<td style="background-color: #F3F3F3">商品管理</td>
		<td style="background-color: #F3F3F3">订单管理</td>
		<td style="background-color: #F3F3F3">会员管理</td>
		<td style="background-color: #F3F3F3">文章管理</td>
		<td style="background-color: #F3F3F3">留言管理</td>
		<td style="background-color: #F3F3F3">评论管理</td>
		<td style="background-color: #F3F3F3">帮助中心</td>
		<td style="background-color: #F3F3F3">权限管理</td>
		<td style="background-color: #F3F3F3">管理人员</td>
		<td rowspan="2" align="center"><a href="admin_info_modi.asp?admin_info_id=<%=admin_info_id%>">修改</a></td>
		<td rowspan="2" align="center"><a href="admin_info_PassWordModiById.asp?admin_info_id=<%=admin_info_id%>">密码修改</a></td>
	</tr>
	<tr>
    <%
	    fla=split(admin_info_flag,",")
        for i=0 to ubound(fla)
    %>
		<td class="altbg2">
		<p align="center">
		<input type="checkbox" name="<%=i%>" value="1" <%if fla(i)=1 then response.write "checked" %> disabled></td>
    <%  next %>
	</tr>
	<%
         rs.movenext
         i=i+1
         wend
    %>
	<tr>
		<td colspan="15">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选 
        <input type="submit" name="action" value="删除" onclick="{if(confirm('删除后将无法恢复，您确定要删除选定的信息吗？')){this.document.form1.submit();return true;}return false;}">&nbsp;
		<input type="button" value="添加" name="action1" onclick="window.location='admin_info_add.asp'"></td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    </form>
</tbody>
</table>
    <%
        call PageControl(iCount,maxpage,page)
    end if
    rs.close
    set rs=nothing
    %>


</body>

</html>
 
