<center><%
'���ϲ�������д��
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp-->
<!--#include file=include/Pages.asp-->
<%
dim count
set rs=server.createobject("adodb.recordset")
sql = "select prod_smallclass_bid,prod_smallclass_id,prod_smallclass_name from prod_smallclass order by prod_smallclass_bid desc"
rs.open sql,conn,1,1
set prod_smallclass_bid =rs(0)
set prod_smallclass_id  =rs(1)
set prod_smallclass_name=rs(2)
%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
subcat[0] = new Array("�˴���������С��","<%=prod_smallclass_bid%>","");
        <%
        count = 1
        do while not rs.eof 
        ss=prod_smallclass_bid
        %>
subcat[<%=count%>] = new Array("<%=prod_smallclass_name%>","<%=prod_smallclass_bid%>","<%=prod_smallclass_id%>");
        <%
        count = count + 1
        rs.movenext
        if prod_smallclass_bid<>ss then
        %>
subcat[<%=count%>] = new Array("�˴���������С��","<%=prod_smallclass_bid%>","");   
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
<%
call up("��Ʒ����","��Ʒ����","��Ʒ����")

'�������в�-����
response.write "			<form name=form1 action=Product_ListSearch.asp method=get>"
response.write "			<tr><td valign=top>��Ʒ���:</td>"
response.write "				<td>"
response.write "					<select size=5 name=bid onChange=changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)>"
response.write "					<option value=''>��ѡ�����</option>"
		    						sql="select prod_bigclass_id,prod_bigclass_name from prod_bigclass order by prod_bigclass_id desc"
		    						set rs=conn.execute (sql)
		    						set prod_bigclass_id=rs(0)
		    						set prod_bigclass_name=rs(1)
		    						do while not rs.eof
		    						response.write "<option value="&prod_bigclass_id&">"&prod_bigclass_name&"</option>"
		    						rs.movenext
		    						loop
		    						rs.close
		    						set rs=nothing
response.write "					</select>"
response.write "					<select name=sid size=5>"
response.write "					<option value=''>��ѡ��С��</option>"
response.write "					</select>"
response.write "		 		</td>"
response.write "			</tr>"
response.write "			<tr><td>��Ʒ���ư���:</td><td><input type=text name=name size=30> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
response.write "			<tr><td>��Ʒ���ݰ���:</td><td><input type=text name=detail size=30></td></tr>"
response.write "			<tr><td>��Ʒ�۸�(��վ��):</td><td><input size=6 name=UserPriceMin>Ԫ(Сֵ)&nbsp; �� <input size=6 name=UserPriceMax>Ԫ(��ֵ)</td></tr>"
response.write "			<tr><td>��Ʒ����:</td><td><input type=checkbox name=flag1 value=1>�Ƽ�&nbsp; <input type=checkbox name=flag2 value=2>��Ʒ <input type=checkbox name=flag value=3>�ؼ�</td></tr>"
response.write "			<tr><td>  </td><td><input class=button type=submit value='��ʼ����'></td></tr>"
response.write "			</form>"

call down()
%></center>