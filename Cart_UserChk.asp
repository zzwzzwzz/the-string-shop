<center>
<center>
<%
dim dbpath,urlpath
dbpath=""
urlpath="Cart_Order.asp"
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<script type="text/javascript">
function chsubmit1()
{
 if (document.form1.LoginName.value == "")        
  {        
    window.alert("�û�������Ϊ�գ�");        
    document.form1.LoginName.focus();        
    return (false);}  
  
        var filter=/^\s*[@.A-Za-z0-9_-]{3,30}\s*$/;
        if (!filter.test(document.form1.LoginName.value)) { 
                window.alert("�û�����д����ȷ,��������д����ʹ�õ��ַ�Ϊ��A-Z a-z 0-9 _ - .)���Ȳ�С��3���ַ���������30���ַ���ע�ⲻҪʹ�ÿո�"); 
                document.form1.LoginName.focus();
                document.form1.LoginName.select();
                return (false); 
                }
 if (document.form1.LoginPass.value == "")        
  {        
    window.alert("���벻��Ϊ�գ�");        
    document.form1.LoginPass.focus();        
    return (false);}  
  
        var filter=/^\s*[.A-Za-z0-9_-]{5,15}\s*$/;
        if (!filter.test(document.form1.LoginPass.value)) { 
                window.alert("������д����ȷ,��������д����ʹ�õ��ַ�Ϊ��A-Z a-z 0-9 _ - .)���Ȳ�С��5���ַ���������15���ַ���ע�ⲻҪʹ�ÿո�"); 
                document.form1.LoginPass.focus();
                document.form1.LoginPass.select();
                return (false); 
                }
 }
</script>
<%
call up("�������ѡ��","�������ѡ��","<a href=cart_list.asp>���ﳵ</a> &raquo; �������ѡ��")

response.write  "<tr><td colspan=2 height=8></td></tr>"&_
				"<tr>"&_
				"	<td width=50% valign=top style='border-right: 1px solid #CCCCCC'>"&_
				"		<table width=100% ><form action=User_LoginCheck.asp method=post name=form1 onsubmit=return chsubmit1();>"&_
				"			<input type=hidden name=urlpath value="&urlpath&">"&_
				"			<tr><td colspan=2>&nbsp;&nbsp;<b>�Ի�Ա��ݽ��㶩����</b></td></tr>"&_
				"			<tr><td>&nbsp;&nbsp;&nbsp;�û���:</td><td><input type=text size=14 name=loginname></td></tr>"&_
				"			<tr><td>&nbsp;&nbsp;&nbsp;�ܡ���:</td><td><input type=password size=14 name=loginpass></td></tr>"&_
				"			<tr><td></td><td><input type=submit value=��¼> <input type=button value=ע�� onclick=window.location='User_Reg.asp?urlpath="&urlpath&"'> <a href=Member_PassWordGet.asp>�������룿</a></td></tr>"&_
				"		</form></table>"&_
				"	</td>"&_
				"	<td width=50% valign=top>"&_
				"		<table width=100% ><tr><td>&nbsp;&nbsp;<b>���ο���ݽ��㶩����</b></td></tr>"&_
				"		<tr><td>&nbsp;&nbsp;<input onclick=document.location.href='Cart_Order.asp'; type=button value=����></td></tr></table>"&_
				"	</td>"&_
				"</tr>"

call down()
%>
</center>
</center>