<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<%
flag=my_request("flag",1)
action=my_request("action",0)
if action="save" then
   call update()
end if

'���̣���������
sub update()
    id=my_request("id",0)
    kucun=my_request("kucun",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '�ж�����news_info_id�й��м�ά
        for v=1 to pp
            id=request("id")(v)
            kucun=request("kucun")(v)
            conn.execute ("update [product_info] set product_info_kucun="&kucun&" where id="&id)
        next
        call ok("���ѳɹ���������Ʒ�������Ϣ��","product_kucun_List.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
<title>��Ʒ���-����</title>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="2" class="header">��Ʒ���-����</td>
	</tr>
	<tr>
		<td colspan="2">
		<p align="center"><font face="����">��</font><a href="Product_KuCun_List.asp">�鿴������Ʒ���</a><font face="����">��</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<font face="����">��</font><a href="?flag=1">�鿴���Ϊ<font color="#FF0000">0</font>����Ʒ</a><font face="����">��</font></td>
	</tr>
	<tr class="altbg1">
		<td>��Ʒ����</td>
		<td>�������</td>
	</tr>
    <form name="form1" action="product_kucun_List.asp" method="post">
    <input name=action value=save type=hidden>
	<%
    set rs=server.createobject("adodb.recordset")
    if flag=1 then
    	sql="select id,product_info_name,product_info_kucun from product_info where product_info_kucun=0 order by id desc"
    else
    	sql="select id,product_info_name,product_info_kucun from product_info order by id desc"
    end if
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=2 align=center>Ŀǰ���������Ʒ�����Ϣ</td></tr>"
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
        while not rs.eof and i<=rs.pagesize
    %>
	<tr>
		<td><input type=hidden name=id value=<%=rs("id")%>><a href=product_info_modi.asp?id=<%=rs("id")%>><%=rs("product_info_name")%></a></td>
		<td>
		<input type=text name=kucun value=<%=rs("product_info_kucun")%> size="4"></td>
	</tr>
	<%
         rs.movenext
         i=i+1
     wend
    %>
	<tr>
		<td>��</td>
		<td><input type="submit" name="b1" value=" �� �� " ></td>
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