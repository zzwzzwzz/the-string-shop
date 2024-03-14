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
<title>����Ա-������Ա��Ϣ-����</title>
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
    admin_info_id=my_request("admin_info_id",0)
    if admin_info_id<>"" then
       pp=ubound(split(admin_info_id,","))+1 '�ж�����id�й��м�ά
       for v=1 to pp
          id=request("admin_info_id")(v)     
          conn.execute ("delete from [admin_info] where admin_info_id="&id)
       next

      call ok("��ѡ��Ϣ�ѳɹ�ɾ����","admin_info_list.asp")
    end if
end sub

%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="15" class="header">������Ա��Ϣ-����</td>
	</tr>	
    <tr>
		<td class="altbg2" colspan="15"></td>
	</tr>
	<tr>
		<td class="altbg1">ѡ��</td>
		<td class="altbg1">��ʵ����</td>
		<td class="altbg1">�û���</td>
		<td class="altbg1" colspan="10">����Ȩ�޷������</td>
		<td class="altbg1" align="center">�޸�</td>
		<td class="altbg1" align="center">�����޸�</td>
	</tr>
	<form name="form1" action="admin_info_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select admin_info_id,admin_info_flag,admin_info_RealName,admin_info_UserName from admin_info order by admin_info_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=15 align=center>Ŀǰ���޹�����Ա��Ϣ,<a href=admin_info_add.asp>������!</a></td></tr>"
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
		<td style="background-color: #F3F3F3">��������</td>
		<td style="background-color: #F3F3F3">��Ʒ����</td>
		<td style="background-color: #F3F3F3">��������</td>
		<td style="background-color: #F3F3F3">��Ա����</td>
		<td style="background-color: #F3F3F3">���¹���</td>
		<td style="background-color: #F3F3F3">���Թ���</td>
		<td style="background-color: #F3F3F3">���۹���</td>
		<td style="background-color: #F3F3F3">��������</td>
		<td style="background-color: #F3F3F3">Ȩ�޹���</td>
		<td style="background-color: #F3F3F3">������Ա</td>
		<td rowspan="2" align="center"><a href="admin_info_modi.asp?admin_info_id=<%=admin_info_id%>">�޸�</a></td>
		<td rowspan="2" align="center"><a href="admin_info_PassWordModiById.asp?admin_info_id=<%=admin_info_id%>">�����޸�</a></td>
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
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
        <input type="submit" name="action" value="ɾ��" onclick="{if(confirm('ɾ�����޷��ָ�����ȷ��Ҫɾ��ѡ������Ϣ��')){this.document.form1.submit();return true;}return false;}">&nbsp;
		<input type="button" value="����" name="action1" onclick="window.location='admin_info_add.asp'"></td>
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