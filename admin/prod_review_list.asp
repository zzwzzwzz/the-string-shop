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
<title>��Ʒ������Ϣ-����</title>
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

'���̣�����ɾ��
sub del()
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '�ж�����news_info_id�й��м�ά
        for v=1 to pp
            id=request("id")(v)
            conn.execute ("delete from [prod_review] where prod_review_id="&id)
        next
        call ok("��ѡ��Ϣ�ѳɹ�ɾ����","prod_review_list.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ʒ-������Ϣ-����</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">��Ʒ������Ϣ-����</td>
	</tr>
    <tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<tr>
		<td class="altbg1">ѡ��</td>
		<td class="altbg1">��Ʒ������Ϣ</td>
	</tr>
	<form name="form1" action="prod_review_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select prod_review_id,prod_review_pid,prod_review_detail,prod_review_name,prod_review_time,prod_review_backdetail,prod_review_BackTime from prod_review order by prod_review_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=2 align=center>Ŀǰ������Ʒ������Ϣ!</a></td></tr>"
    else
        rs.PageSize =10 'ÿҳ��¼����
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
        
        while not rs.eof and i<=rs.pagesize
            set prod_review_id        =rs(0)
            set prod_review_pid       =rs(1)
            set prod_review_detail    =rs(2)
            set prod_review_name      =rs(3)
            set prod_review_time      =rs(4)
            set prod_review_backdetail=rs(5)
            set prod_review_BackTime  =rs(6)
            sql1="select product_info_name from product_info where id="&prod_review_pid
            set rs1=conn.execute (sql1)
            prod_info_name=rs1(0)
            rs1.close
            set rs1=nothing  
    %>
	<tr>
		<td valign="top"><input type="checkbox" name="id" value="<%=prod_review_id%>"></td>
		<td valign="top">������Ʒ��<b><a href=../product_detail.asp?id=<%=prod_review_pid%> target="_blank"><%=prod_info_name%></a></b><br>
		�������ݣ�<%=prod_review_detail%><br>
        <font color="#808080">�û�����</font><font color="#808080"><%=prod_review_name%></font><br>
        <font color="#808080">����ʱ�䣺</font><font color="#808080"><%=prod_review_time%></font><br>
		<hr color="#CCCCCC" size="1">
		<%if prod_review_backdetail<>"" then%><font color="#cccccc"><b>�ѻظ���</b></font><%=prod_review_backdetail%><font color="#999999">( �ظ�ʱ�䣺<%=prod_review_BackTime%> )</font><input type="button" value="�༭�ظ�" name="action1" onclick="window.location='prod_review_back.asp?prod_review_id=<%=prod_review_id%>'"><%else%><input type="button" value="�ظ�" name="action1" onclick="window.location='prod_review_back.asp?prod_review_id=<%=prod_review_id%>'"><%end if%>
        </td>
	</tr>
	<%
         rs.movenext
         i=i+1
         wend
    %>
	<tr>
		<td colspan="2">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
        <input type="submit" name="action" value="ɾ��" onclick="{if(confirm('ɾ�����޷��ָ�����ȷ��Ҫɾ��ѡ������Ϣ��')){this.document.form1.submit();return true;}return false;}"></td>
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
 
