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
<title>������Ϣ-����</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language = "JavaScript">   
//ȫѡ����    
function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }

</script>
<%
action=my_request("action",0)
if action="ɾ��" then
    call del()
end if

'���̣������������վ����ע��ɾ�����
sub del()
    order_info_id=my_request("order_info_id",0)
    if order_info_id<>"" then
       pp=ubound(split(order_info_id,","))+1 '�ж�����id�й��м�ά
       for v=1 to pp
          id=request("order_info_id")(v)     
          conn.execute ("update [order_info] set order_info_recycle=1 where order_info_id="&id)
       next

      call ok("��ѡ��Ϣ�Ѽ������վ����ע��ɾ����ǣ�","order_info_list.asp")
    end if
end sub

%>

</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="6" class="header">��������</td>
	</tr>
    <tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
    <tr>
		<td class="altbg1">ѡ��</td>
		<td class="altbg1">������</td>
		<td class="altbg1">���</td>
		<td class="altbg1">�ջ�������</td>
		<td class="altbg1">�µ�ʱ��</td>
		<td class="altbg1">����״̬</td>
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
        response.write "<tr><td colspan=7 align=center>Ŀǰ���޶�����Ϣ</td></tr>"
    else
        rs.PageSize =20 'ÿҳ��¼����
        iCount=rs.RecordCount '��¼����
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
                order_info_CheckStates="�¶���(δȷ��)"
            case "1"
                order_info_CheckStates="�˿�����ȡ��"
            case "2"
                order_info_CheckStates="��Ч������ȡ��"
            case "3"
                order_info_CheckStates="��ȷ�ϣ�������"
            case "4"
                order_info_CheckStates="�ѷ��������ջ�"
            case "5"
                order_info_CheckStates="����֧���ɹ�"
            case "6"
                order_info_CheckStates="�������"
        end select           
    %>
	<tr>
		<td><input type="checkbox" name="order_info_id" value="<%=order_info_id%>"></td>
		<td><a href=order_info_Modi.asp?order_info_id=<%=order_info_id%>><%=order_info_no%></a></td>
		<td><%=order_info_AllCost%>Ԫ</td>
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
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
        <input type="submit" name="action" value="ɾ��" onclick="{if(confirm('��ʾ����ȷ��Ҫɾ����ѡ���Ķ�����')){this.document.form1.submit();return true;}return false;}">&nbsp;
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