<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=8
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<%
action=my_request("action",0)
if action="setdefault" then
    call setdefault()
end if

sub setdefault()
    root_info_skin=my_request("skin",0)
    if root_info_skin="" then
        response.redirect "error.htm"
        response.end
    else
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select top 1 * from root_info"
        rs.open sql,conn,1,3
        rs("root_info_skin")=root_info_skin
        rs.update
        rs.close
        set rs=nothing

        call ok("���ѳɹ����õ�ǰʹ��ģ����Ϣ��","root_model_list.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��վģ��-����</title>
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
       pp=ubound(split(id,","))+1 '�ж�����id�й��м�ά
       for v=1 to pp
          id=request("id")(v)
          conn.execute ("delete from [root_model] where id="&id)
       next
       call ok("��ѡ��Ϣ�ѳɹ�ɾ����","root_model_list.asp")
    end if
end sub

Function CheckDir(FolderPath)
    folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
        CheckDir = True
    Else
        CheckDir = False
    End if
    Set fso1 = nothing
End Function
%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="6" class="header">��վģ��-����</td>
	</tr>
	<tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<tr>
		<td class="altbg1">ѡ��</td>
		<td class="altbg1">ģ������</td>
		<td class="altbg1">��Ӧ��ʽ���ļ���</td>
		<td class="altbg1">�Ƿ�ǰʹ��ģ��</td>
		<td class="altbg1" align="center">�޸�</td>
	</tr>
    <form name="form1" action="Root_Model_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select root_model_name,root_model_css,id from root_model order by id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=6 align=center>Ŀǰ������վģ����Ϣ,<a href=root_model_add.asp>������!</a></td></tr>"
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
        
        set root_model_name=rs(0)
        set root_model_css=rs(1)
        set id=rs(2)

        while not rs.eof and i<=rs.pagesize
        
			Set rs1=Server.CreateObject("ADODB.Recordset")
			sql1="select root_info_skin from root_info where id=1"
			rs1.open sql1,conn,1,1
			root_info_skin=rs1(0)
			rs1.close
			set rs1=nothing
			if root_model_css=root_info_skin then
				txtcss="<font color=red>��</font>"
			else 
				txtcss="<a href=?action=setdefault&skin="&root_model_css&">��(�������)</a>"
			end if
    %>
	<tr>
		<td><input type="checkbox" name="id" value="<%=id%>"></td>
		<td><%=root_model_name%></td>
		<td><%=root_model_css%>.css</td>
		<td><%=txtcss%></td>
		<td align="center"><a href="root_model_modi.asp?id=<%=id%>">�޸�</a></td>
	</tr>
	<%
         rs.movenext
         i=i+1
     wend
    %>
	<tr>
		<td colspan="6">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
        <input type="submit" name="action" value="ɾ��" onClick="{if(confirm('ɾ�����޷��ָ�����ȷ��Ҫɾ��ѡ������Ϣ��')){this.document.form1.submit();return true;}return false;}">&nbsp;
		<input type="button" value="����" name="action1" onClick="window.location='root_model_add.asp'"></td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    <%
        call PageControl(iCount,maxpage,page)
    end if
    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
    %>
    </form>
</tbody>
</table>


</body>

</html>

