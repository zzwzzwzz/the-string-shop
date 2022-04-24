<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=5
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品评论信息-管理</title>
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
        pp=ubound(split(id,","))+1 '判断数组news_info_id中共有几维
        for v=1 to pp
            id=request("id")(v)
            conn.execute ("delete from [prod_review] where prod_review_id="&id)
        next
        call ok("所选信息已成功删除！","prod_review_list.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品-评论信息-管理</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">商品评论信息-管理</td>
	</tr>
	<tr>
		<td class="altbg1">选中</td>
		<td class="altbg1">商品评论信息</td>
	</tr>
	<form name="form1" action="prod_review_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select prod_review_id,prod_review_pid,prod_review_detail,prod_review_name,prod_review_time,prod_review_IP,prod_review_backdetail,prod_review_BackTime from prod_review order by prod_review_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=2 align=center>目前暂无商品评论信息!</a></td></tr>"
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
            set prod_review_id        =rs(0)
            set prod_review_pid       =rs(1)
            set prod_review_detail    =rs(2)
            set prod_review_name      =rs(3)
            set prod_review_time      =rs(4)
            set prod_review_IP        =rs(5)
            set prod_review_backdetail=rs(6)
            set prod_review_BackTime  =rs(7)
            sql1="select product_info_name from product_info where id="&prod_review_pid
            set rs1=conn.execute (sql1)
            prod_info_name=rs1(0)
            rs1.close
            set rs1=nothing  
    %>
	<tr>
		<td valign="top"><input type="checkbox" name="id" value="<%=prod_review_id%>"></td>
		<td valign="top">评论商品：<b><a href=../product_detail.asp?id=<%=prod_review_pid%> target="_blank"><%=prod_info_name%></a></b><br>
		评论内容：<%=prod_review_detail%><br>
        <font color="#808080">用户名：</font><font color="#808080"><%=prod_review_name%></font><br>
        <font color="#808080">评论时间：</font><font color="#808080"><%=prod_review_time%></font><br>
		<hr color="#CCCCCC" size="1">
		<%if prod_review_backdetail<>"" then%><font color="#cccccc"><b>已回复：</b></font><%=prod_review_backdetail%><font color="#999999">( 回复时间：<%=prod_review_BackTime%> )</font><input type="button" value="编辑回复" name="action1" onclick="window.location='prod_review_back.asp?prod_review_id=<%=prod_review_id%>'"><%else%><input type="button" value="回复" name="action1" onclick="window.location='prod_review_back.asp?prod_review_id=<%=prod_review_id%>'"><%end if%>
        </td>
	</tr>
	<%
         rs.movenext
         i=i+1
         wend
    %>
	<tr>
		<td colspan="2">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选 
        <input type="submit" name="action" value="删除" onclick="{if(confirm('删除后将无法恢复，您确定要删除选定的信息吗？')){this.document.form1.submit();return true;}return false;}"></td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    </form>
</tbody>
</table>
    <%
        call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
    end if
    rs.close
    set rs=nothing
    %>

</body>

</html>
 
