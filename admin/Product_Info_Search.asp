<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ʒ�߼�����</title>
<link rel="stylesheet" type="text/css" href="style.css">
<%
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from prod_smallclass order by prod_smallclass_bid desc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
subcat[0] = new Array("�˴���������С��","<%= trim(rs("prod_smallclass_bid"))%>","");
        <%
        count = 1
        do while not rs.eof 
        ss=trim(rs("prod_smallclass_bid"))
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("prod_smallclass_name"))%>","<%= trim(rs("prod_smallclass_bid"))%>","<%= trim(rs("prod_smallclass_id"))%>");
        <%
        count = count + 1
        rs.movenext
        if trim(rs("prod_smallclass_bid"))<>ss then
        %>
subcat[<%=count%>] = new Array("�˴���������С��","<%= trim(rs("prod_smallclass_bid"))%>","");   
        <%
        count = count + 1
        end if
        loop
        rs.close
        set rs=nothing
        %>
onecount=<%=count%>;

//����л�
function changelocation(locationid)
    {
    document.form1.sid.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form1.sid.options[document.form1.sid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    } 
</script>
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Product_Info_List.asp" method="get" name="form1">
    <tr>
		<td colspan="2" class="title">��Ʒ�߼�����</td>
	</tr>
	<tr>
		<td>��Ʒ���ƣ�</td>
		<td><input type="text" name="KeyWord" size="20"></td>
	</tr>
	<tr>
		<td>�������</td>
		<td>
			<select name="bid" onChange="changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)">
		    <option value="">��ѡ�����</option>
		    <%
		    sql="select * from prod_bigclass order by prod_bigclass_id desc"
		    set rs=conn.execute (sql)
		    do while not rs.eof
		    %>
		    <option value="<%=rs("prod_bigclass_id")%>"><%=rs("prod_bigclass_name")%></option>
		    <%
		    rs.movenext
		    loop
		    rs.close
		    set rs=nothing
		    %>
            </select><select name="sid"> 
            <option value="">��ѡ��С��</option>
            </select>
        </td>
	</tr>
	<tr>
		<td>��Ʒ���ݺ���</td>
		<td><input type="text" name="product_info_Detail" size="20"></td>
	</tr>
	<tr>
		<td>��վ�۸�Χ��</td>
		<td><input type="text" name="prod_info_PriceSMin" size="6">Ԫ(Сֵ)&nbsp; �� 
		<input type="text" name="prod_info_PriceSMax" size="6">Ԫ(��ֵ)</td>
	</tr>
	<tr>
		<td>�����������</td>
		<td><input type="radio" CHECKED value="1" name="sort">ʱ�併 
		<input type="radio" value="2" name="sort">ʱ���� 
		<input type="radio" value="3" name="sort">��Ž� 
		<input type="radio" value="4" name="sort">����� 
		<input type="radio" value="5" name="sort">��Ʒ���� 
		<input type="radio" value="6" name="sort">�������</td>
	</tr>
	<tr>
		<td>��</td>
		<td><input type="submit" value="  ��  ��  " name="B1">&nbsp;&nbsp;&nbsp;
			<input type="reset" value="  ��  ��  " name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>

