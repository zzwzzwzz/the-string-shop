<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=4
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

//���̣�����ɾ��
sub del()
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '�ж�����news_info_id�й��м�ά
        for v=1 to pp
            id=request("id")(v)
            conn.execute ("delete from [news_info] where id="&id)
        next
        call ok("��ѡ��Ϣ�ѳɹ�ɾ����","News_Info_List.asp")
    end if
end sub
%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="5" class="header">������Ϣ-����</td>
	</tr>
	<tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<tr class="altbg1">
		<td>ѡ��</td>
		<td>��Ϣ����</td>
		<td>����ʱ��</td>
		<td>�����/��</td>
		<td>�༭</td>
	</tr>
	<form name="form1" action="News_Info_List.asp" method="post">
    <%
    set rs=server.createobject("adodb.recordset")
    sql="select id,news_info_title,news_info_addtime,news_info_hitnums from news_info order by id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=5 align=center>�������������Ϣ,<a href=News_Info_Add.asp>������!</a></td></tr>"
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
      	
      	set id                = rs(0)
      	set news_info_title   = rs(1)
      	set news_info_addtime = rs(2)
      	set news_info_hitnums = rs(3)
      	
      	while not rs.eof and i<=rs.pagesize
    %>
	<tr>
		<td><input type="checkbox" name="id" value="<%=id%>">   </td>
		<td><a href=News_Info_Modi.asp?id=<%=id%>><%=news_info_title%></a></td>
		<td><%=datevalue(news_info_addtime)%>                             </td>
		<td><%=news_info_hitnums%>                                        </td>
		<td><a href=News_Info_Modi.asp?id=<%=id%>>�༭</a>                </td>
	</tr>
	<%
        rs.movenext
        i=i+1
        wend
    %>
	<tr>
		<td colspan="5">
		<input type="checkbox" name="chkall" onclick="CheckAll(this.form)">ȫѡ 
        <input type="submit" name="action" value="ɾ��" onclick="{if(confirm('ɾ�����޷��ָ�����ȷ��Ҫɾ��ѡ������Ϣ��')){this.document.form1.submit();return true;}return false;}">&nbsp;
        <input type="button" value="������Ϣ-����" name="action1" onclick="window.location='News_Info_Add.asp'"></td>
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