<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<script type="text/javascript">
function chsubmit()
{
 if (document.form_1.LoginName.value == "")        
  {        
    window.alert("�û�������Ϊ�գ�");        
    document.form_1.LoginName.focus();        
    return (false);}  
  
        var filter=/^\s*[@.A-Za-z0-9_-]{3,30}\s*$/;
        if (!filter.test(document.form_1.LoginName.value)) { 
                window.alert("�û�����д����ȷ,��������д����ʹ�õ��ַ�Ϊ��A-Z a-z 0-9 _ - .)���Ȳ�С��3���ַ���������30���ַ���ע�ⲻҪʹ�ÿո�"); 
                document.form_1.LoginName.focus();
                document.form_1.LoginName.select();
                return (false); 
                }
 if (document.form_1.LoginPass.value == "")        
  {        
    window.alert("���벻��Ϊ�գ�");        
    document.form_1.LoginPass.focus();        
    return (false);}  
  
        var filter=/^\s*[.A-Za-z0-9_-]{5,15}\s*$/;
        if (!filter.test(document.form_1.LoginPass.value)) { 
                window.alert("������д����ȷ,��������д����ʹ�õ��ַ�Ϊ��A-Z a-z 0-9 _ - .)���Ȳ�С��5���ַ���������15���ַ���ע�ⲻҪʹ�ÿո�"); 
                document.form_1.LoginPass.focus();
                document.form_1.LoginPass.select();
                return (false); 
                }
 }

</script>
<%
urlpath=my_request("urlpath",0)

call up("��Ա��½","��Ա��½","��Ա��½")

response.write  "<form name=form_1 action=User_loginCheck.asp method=post onsubmit='return chsubmit();'>"&_
				"<input type=hidden name=urlpath value="&urlpath&">"&_
				"	<tr><td colspan=2 align=center height=40><b>����д�û��������룺</b></td></tr>"&_
				"	<tr><td align=right width=40% >&nbsp;�û���:</td><td><input type=text size=14 name=LoginName></td></tr>"&_
				"	<tr><td align=right width=40% >&nbsp;�ܡ���:</td><td><input type=password size=14 name=LoginPass></td></tr>"&_
				"	<tr><td width=40% ></td><td><input class=button type=submit value=' �� ¼ '>    <input class=button type=button value=' ע �� 'onclick=window.location='User_Reg.asp?urlpath="&urlpath&"'> <a href=User_PassWordGet.asp>��������</a></td></tr>"&_
				"</form>"
call down()
%></center>