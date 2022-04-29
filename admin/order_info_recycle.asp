<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=2
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>订单-订单信息-管理</title>
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
if action="彻底删除" then
    call del()
elseif action="恢复订单" then
    call restore()
end if

//过程：批量彻底删除
sub del()
    order_info_id=my_request("order_info_id",0)
    if order_info_id<>"" then
        pp=ubound(split(order_info_id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("order_info_id")(v)     
            conn.execute ("delete from [order_info] where order_info_id="&id)
        next

        call ok("所选信息已成功删除！","order_info_recycle.asp")
    end if
end sub

//过程：批量恢复订单
sub restore()
    order_info_id=my_request("order_info_id",0)
    if order_info_id<>"" then
        pp=ubound(split(order_info_id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("order_info_id")(v)     
            conn.execute ("update [order_info] set order_info_recycle=0 where order_info_id="&id)
        next

        call ok("所选订单已成功恢复！","order_info_recycle.asp")
    end if
end sub
%>

</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="7" class="header">订单回收站</td>
	</tr>
    <tr>
		<td class="altbg2" colspan="7"></td>
	</tr>
    <tr>
		<td class="altbg1">选中</td>
		<td class="altbg1">订单号</td>
		<td class="altbg1">金额</td>
		<td class="altbg1">会员ID</td>
		<td class="altbg1">收货人姓名</td>
		<td class="altbg1">下单时间</td>
		<td class="altbg1">订单状态</td>
	</tr>
	<form name="form1" action="order_info_recycle.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select order_info_id,order_info_no,order_info_AllCost,order_info_UserName,order_info_RealName,order_info_BuyTime,order_info_CheckStates from order_info where order_info_recycle=1 order by order_info_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=7 align=center>订单回收站为空</td></tr>"
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
        set order_info_id         =rs(0)
        set order_info_no         =rs(1)
        set order_info_AllCost    =rs(2)
        set order_info_UserName   =rs(3)
        set order_info_RealName   =rs(4)
        set order_info_BuyTime    =rs(5)
        set order_info_CheckStates=rs(6)
        while not rs.eof and i<=rs.pagesize
        select case order_info_CheckStates
            case 0
                order_info_CheckStates="新订单(未确认)"
            case 1
                order_info_CheckStates="会员自行取消"
            case 2
                order_info_CheckStates="无效单，已取消"
            case 3
                order_info_CheckStates="已确认，待付款"
            case 4
                order_info_CheckStates="已发货，待收货"
            case 5
                order_info_CheckStates="在线支付成功"
            case 6
                order_info_CheckStates="订单完成"
        end select           
    %>
	<tr>
		<td><input type="checkbox" name="order_info_id" value="<%=order_info_id%>"></td>
		<td><a href=order_info_Modi.asp?order_info_id=<%=order_info_id%>><%=order_info_no%></a></td>
		<td><%=order_info_AllCost%>元</td>
		<td><%=order_info_UserName%></td>
		<td><%=order_info_RealName%></td>
		<td><%=order_info_BuyTime%></td>
		<td><%=order_info_CheckStates%></td>
	</tr>
	<%
         rs.movenext
         i=i+1
         wend
    %>
	<tr>
		<td colspan="7">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选 
        <input type="submit" name="action" value="彻底删除" onclick="{if(confirm('提示：删除后将无法恢复，您确定要删除选定的订单吗？')){this.document.form1.submit();return true;}return false;}">&nbsp;
	    <input type="submit" name="action" value="恢复订单" onclick="{if(confirm('提示：您确定要恢复所选定的订单吗？')){this.document.form1.submit();return true;}return false;}"></td>
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
 
