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
<title>����-������Ϣ-����</title>
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
if action="����ɾ��" then
    call del()
elseif action="�ָ�����" then
    call restore()
end if

'���̣���������ɾ��
sub del()
    order_info_id=my_request("order_info_id",0)
    if order_info_id<>"" then
        pp=ubound(split(order_info_id,","))+1 '�ж�����id�й��м�ά
        for v=1 to pp
            id=request("order_info_id")(v)     
            conn.execute ("delete from [order_info] where order_info_id="&id)
        next

        call ok("��ѡ��Ϣ�ѳɹ�ɾ����","order_info_recycle.asp")
    end if
end sub

'���̣������ָ�����
sub restore()
    order_info_id=my_request("order_info_id",0)
    if order_info_id<>"" then
        pp=ubound(split(order_info_id,","))+1 '�ж�����id�й��м�ά
        for v=1 to pp
            id=request("order_info_id")(v)     
            conn.execute ("update [order_info] set order_info_recycle=0 where order_info_id="&id)
        next

        call ok("��ѡ�����ѳɹ��ָ���","order_info_recycle.asp")
    end if
end sub
%>

</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="7" class="header">��������վ</td>
	</tr>
    <tr>
		<td class="altbg2" colspan="7"></td>
	</tr>
    <tr>
		<td class="altbg1">ѡ��</td>
		<td class="altbg1">������</td>
		<td class="altbg1">���</td>
		<td class="altbg1">��ԱID</td>
		<td class="altbg1">�ջ�������</td>
		<td class="altbg1">�µ�ʱ��</td>
		<td class="altbg1">����״̬</td>
	</tr>
	<form name="form1" action="order_info_recycle.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select order_info_id,order_info_no,order_info_AllCost,order_info_UserName,order_info_RealName,order_info_BuyTime,order_info_CheckStates from order_info where order_info_recycle=1 order by order_info_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=7 align=center>��������վΪ��</td></tr>"
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
        set order_info_UserName   =rs(3)
        set order_info_RealName   =rs(4)
        set order_info_BuyTime    =rs(5)
        set order_info_CheckStates=rs(6)
        while not rs.eof and i<=rs.pagesize
        select case order_info_CheckStates
            case 0
                order_info_CheckStates="�¶���(δȷ��)"
            case 1
                order_info_CheckStates="��Ա����ȡ��"
            case 2
                order_info_CheckStates="��Ч������ȡ��"
            case 3
                order_info_CheckStates="��ȷ�ϣ�������"
            case 4
                order_info_CheckStates="�ѷ��������ջ�"
            case 5
                order_info_CheckStates="����֧���ɹ�"
            case 6
                order_info_CheckStates="�������"
        end select           
    %>
	<tr>
		<td><input type="checkbox" name="order_info_id" value="<%=order_info_id%>"></td>
		<td><a href=order_info_Modi.asp?order_info_id=<%=order_info_id%>><%=order_info_no%></a></td>
		<td><%=order_info_AllCost%>Ԫ</td>
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
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
        <input type="submit" name="action" value="����ɾ��" onclick="{if(confirm('��ʾ��ɾ�����޷��ָ�����ȷ��Ҫɾ��ѡ���Ķ�����')){this.document.form1.submit();return true;}return false;}">&nbsp;
	    <input type="submit" name="action" value="�ָ�����" onclick="{if(confirm('��ʾ����ȷ��Ҫ�ָ���ѡ���Ķ�����')){this.document.form1.submit();return true;}return false;}"></td>
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