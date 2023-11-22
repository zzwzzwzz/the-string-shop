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

'���̣�����ɾ��
sub del()
    guest_info_id=my_request("guest_info_id",0)
    if guest_info_id<>"" then
       pp=ubound(split(guest_info_id,","))+1 '�ж�����id�й��м�ά
       for v=1 to pp
          id=request("guest_info_id")(v)     
          conn.execute ("delete from [guest_info] where guest_info_id="&id)
       next

      call ok("��ѡ��Ϣ�ѳɹ�ɾ����","GB_Info_List.asp")
    end if
end sub
%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">������Ϣ-����</td>
	</tr>
    <tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<tr>
		<td class="altbg1">ѡ��</td>
		<td class="altbg1">������������</td>
	</tr>
	<form name="form1" action="GB_Info_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select * from guest_info order by guest_info_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=2 align=center>Ŀǰ����������Ϣ!</a></td></tr>"
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
    %>
	<tr>
		<td valign="top"><input type="checkbox" name="guest_info_id" value="<%=rs("guest_info_id")%>"></td>
		<td valign="top"><font color="#808080">�û�����</font><b><font color="#808080"><%=rs("guest_info_name")%></font></b><br>
		<font color="#808080">Email��</font><font color="#808080"><%=rs("guest_info_email")%></font><br>
		<font color="#808080">����ʱ�䣺</font><font color="#808080"><%=rs("guest_info_time")%></font><br>
		<font color="#808080">�������ݣ�</font><font color="#808080"><%=rs("guest_info_detail")%></font><hr>
		<%if rs("guest_info_backdetail")<>"" then%><font color="#cccccc"><b>�ѻظ���</b></font><font color="#999999"><%=rs("guest_info_backDetail")%> </font>
		<font color="#999999">(�ظ�ʱ�䣺<%=rs("guest_info_BackTime")%>)</font><input type="button" value="�༭�ظ�" name="action1" onclick="window.location='GB_info_back.asp?guest_info_id=<%=rs("guest_info_id")%>'"><%else%><input type="button" value="�ظ�" name="action1" onclick="window.location='GB_info_back.asp?guest_info_id=<%=rs("guest_info_id")%>'"><%end if%>
        </td>
	</tr>
	<%   rs.movenext
         i=i+1
         wend
    %>
	<tr>
		<td colspan="2">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
        <input type="submit" name="action" value="ɾ��" onclick="{if(confirm('ɾ�����޷��ָ�����ȷ��Ҫɾ��ѡ������Ϣ��')){this.document.form1.submit();return true;}return false;}"></td>
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