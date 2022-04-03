<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=2
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<%
search_order_CheckStates=my_request("search_order_CheckStates",0)
search_order_no         =my_request("search_order_no",0)
search_order_RealName   =my_request("search_order_RealName",0)
search_order_email      =my_request("search_order_email",0)
search_order_tel        =my_request("search_order_tel",0)
search_order_mobile     =my_request("search_order_mobile",0)
search_order_address    =my_request("search_order_address",0)
search_order_zip        =my_request("search_order_zip",0)
search_order_BuyTime    =my_request("search_order_BuyTime",0)

Search=""

if search_order_CheckStates<>"" then
    Search=Search & "and order_info_CheckStates="&search_order_CheckStates
end if

if search_order_no<>"" then
    Search=Search & "and order_info_no='"&search_order_no&"'"
end if

if search_order_RealName<>"" then
    Search=Search & "and order_info_RealName = '"&search_order_RealName&"'"
end if

if search_order_email<>"" then
    Search=Search & "and order_info_email = '"&search_order_email&"'"
end if

if search_order_tel<>"" then
    Search=Search & "and order_info_tel = '"&search_order_tel&"'"
end if

if search_order_mobile<>"" then
    Search=Search & "and order_info_mobile = '"&search_order_mobile&"'"
end if

if search_order_address<>"" then
    Search=Search & "and order_info_address like  '%"&search_order_address&"%'"
end if

if search_order_zip<>"" then
    Search=Search & "and order_info_zip = '"&search_order_zip&"'"
end if

if search_order_BuyTime<>"" then
    select case search_order_BuyTime
        case 1   
            DayFrom=dateadd("y",-1,now)
            DayFrom=cdate(DayFrom)
            DayTo=now
            DayTo=cdate(DayTo)
            DayFrom="#"&DayFrom&"#"
            DayTo="#"&DayTo&"#"
            Search=Search & "and order_info_BuyTime Between "&DayFrom&" and "&DayTo&""
        case 2   
            DayFrom=dateadd("y",-2,now)
            DayFrom=cdate(DayFrom)
            DayTo=now
            DayTo=cdate(DayTo)
            DayFrom="#"&DayFrom&"#"
            DayTo="#"&DayTo&"#"
            Search=Search & "and order_info_BuyTime Between "&DayFrom&" and "&DayTo&""
        case 7   
            DayFrom=dateadd("y",-7,now)
            DayFrom=cdate(DayFrom)
            DayTo=now
            DayTo=cdate(DayTo)
            DayFrom="#"&DayFrom&"#"
            DayTo="#"&DayTo&"#"
            Search=Search & "and order_info_BuyTime Between "&DayFrom&" and "&DayTo&""
        case 30   
            DayFrom=dateadd("y",-30,now)
            DayFrom=cdate(DayFrom)
            DayTo=now
            DayTo=cdate(DayTo)
            DayFrom="#"&DayFrom&"#"
            DayTo="#"&DayTo&"#"
            Search=Search & "and order_info_BuyTime Between "&DayFrom&" and "&DayTo&""
    end select
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>订单信息-管理</title>
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

//过程：批量加入回收站，加注了删除标记
sub del()
    order_info_id=my_request("order_info_id",0)
    if order_info_id<>"" then
       pp=ubound(split(order_info_id,","))+1 '判断数组id中共有几维
       for v=1 to pp
          id=request("order_info_id")(v)     
          conn.execute ("update [order_info] set order_info_recycle=1 where order_info_id="&id)
       next

      call ok("所选信息已加入回收站，加注了删除标记！","order_info_list.asp")
    end if
end sub

%>

</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="6" class="header">订单管理</td>
	</tr>
    <tr>
		<td class="altbg1">选中</td>
		<td class="altbg1">订单号</td>
		<td class="altbg1">金额</td>
		<td class="altbg1">收货人姓名</td>
		<td class="altbg1">下单时间</td>
		<td class="altbg1">订单状态</td>
	</tr>
	<form name="form1" action="order_info_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    if search<>"" then
        sql="select order_info_id,order_info_no,order_info_AllCost,order_info_RealName,order_info_BuyTime,order_info_CheckStates from order_info where 1=1 "&Search&" order by order_info_id desc"     
        'response.write sql
        'response.end
    else
        sql="select order_info_id,order_info_no,order_info_AllCost,order_info_RealName,order_info_BuyTime,order_info_CheckStates from order_info order by order_info_id desc"
    end if
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=7 align=center>目前暂无订单信息</td></tr>"
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
        set order_info_RealName   =rs(3)
        set order_info_BuyTime    =rs(4)

        while not rs.eof and i<=rs.pagesize
        order_info_CheckStates    =rs(5)
        select case order_info_CheckStates
            case "0"
                order_info_CheckStates="新订单(未确认)"
            case "1"
                order_info_CheckStates="顾客自行取消"
            case "2"
                order_info_CheckStates="无效单，已取消"
            case "3"
                order_info_CheckStates="已确认，待付款"
            case "4"
                order_info_CheckStates="已发货，待收货"
            case "5"
                order_info_CheckStates="在线支付成功"
            case "6"
                order_info_CheckStates="订单完成"
        end select           
    %>
	<tr>
		<td><input type="checkbox" name="order_info_id" value="<%=order_info_id%>"></td>
		<td><a href=order_info_Modi.asp?order_info_id=<%=order_info_id%>><%=order_info_no%></a></td>
		<td><%=order_info_AllCost%>元</td>
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
		<td colspan="6">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选 
        <input type="submit" name="action" value="删除" onclick="{if(confirm('提示：你确定要删除所选定的订单吗？')){this.document.form1.submit();return true;}return false;}">&nbsp;
	    </td>
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

