<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=3
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ա-��Ա��Ϣ-����</title>
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

//���̣�����ɾ��
sub del()
    user_info_id=my_request("user_info_id",0)
    if user_info_id<>"" then
       pp=ubound(split(user_info_id,","))+1 '�ж�����id�й��м�ά
       for v=1 to pp
          id=request("user_info_id")(v)     
          conn.execute ("delete from [user_info] where user_info_id="&id)
       next

      call ok("��ѡ��Ϣ�ѳɹ�ɾ����","user_info_list.asp")
    end if
end sub

%>

</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="9" class="header">��Ա��Ϣ-����</td>
	</tr>
    <tr>
		<td class="altbg2" colspan="9"></td>
	</tr>
	<tr>
		<td class="altbg1">ѡ��</td>
		<td class="altbg1">��ԱID</td>
		<td class="altbg1">��Ա����</td>
		<td class="altbg1">ע������</td>
		<td class="altbg1">�ϴε�½</td>
		<td class="altbg1">��½����</td>
		<td class="altbg1">״̬</td>
		<td class="altbg1" align="center">�޸�</td>
	</tr>
	<form name="form1" action="user_info_list.asp" method="post">
<%
    set rs=server.createobject("adodb.recordset")
    sql="select user_info_id,user_info_UserName,user_info_RealName,user_info_RegTime,user_info_LastLoginTime,user_info_LoginNums,user_info_states from user_info order by user_info_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=9 align=center>Ŀǰ���޻�Ա��Ϣ!</a></td></tr>"
    else
        rs.PageSize =500 'ÿҳ��¼����
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
        
        set user_info_id=rs(0)
        set user_info_UserName=rs(1)
        set user_info_RealName=rs(2)
        set user_info_RegTime=rs(3)
        set user_info_LastLoginTime=rs(4)
        set user_info_LoginNums=rs(5)
        set user_info_states=rs(6)
        while not rs.eof and i<=rs.pagesize
%>
	<tr>
		<td><input type="checkbox" name="user_info_id" value="<%=user_info_id%>"></td>
		<td><a href=user_info_modi.asp?user_info_id=<%=user_info_id%>><%=user_info_UserName%></a></td>
		<td><%=user_info_RealName%></td>
		<td><%=user_info_RegTime%></td>
		<td><%=user_info_LastLoginTime%></td>
		<td><%=user_info_LoginNums%></td>
		<td><%if user_info_states=0 then response.write "����/ͨ�����" else response.write "<font color=#C0C0C0>����/�����</font>"%></td>
		<td align="center"><a href=user_info_modi.asp?user_info_id=<%=user_info_id%>>�޸�</a></td>
	</tr>
	<%
         rs.movenext
         i=i+1
         wend
    %>
	<tr>
		<td colspan="9">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
        <input type="submit" name="action" value="ɾ��" onclick="{if(confirm('ɾ�����޷��ָ�����ȷ��Ҫɾ��ѡ������Ϣ��')){this.document.form1.submit();return true;}return false;}"></td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    </form>
</tbody>
</table>
    <p align="center"><font color="#C0C0C0">
    <%
        call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
    end if
    rs.close
    set rs=nothing
    %>
</font></p>

</body>

</html>
