<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
call up("����֧��","����֧��","����֧��")
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_info_sitename from root_info where id=1"
rs.open sql,conn,1,1
root_info_sitename=rs(0)
rs.close
set rs=nothing

response.write  "<form name=formorder method=post action=OnlyOne_ByAlipay.asp>"&_
				"<input type=hidden name=product_info_name value="&root_info_sitename&"����֧��������>"&_
				"<tr><td></td><td><img src=images/logo_alipay.gif align=absmiddle><b>����֧������</b></td></tr>"&_
				"<tr><td>������֧����</td><td><input type=text size=30 name=product_info_PriceS> </td></tr>"&_
				"<tr><td>��������ʵ������</td><td><input type=text size=30 name=order_info_realname> </td></tr>"&_
				"<tr><td>��������ϵ�绰��</td><td><input type=text size=30 name=order_info_mobile>	     </td></tr>"&_
				"<tr><td>�����˵������䣺</td><td><input type=text size=30 name=order_info_email>    </td></tr>"&_
				"<tr><td>				 </td><td><input type=submit value=  ��ʼ֧�� >		         </td></tr>"&_
				"</form>"
call down()
%></center>