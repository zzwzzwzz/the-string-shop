<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file="include/md5.asp"-->
<!--#include file=Sub.asp -->
<%
urlpath=my_request("urlpath",0)

action=my_request("action",0)
if action="save" then
    call User_RegSave()
end if

call up("ע���Ա","ע���Ա","ע���Ա")
response.write  "<form name=form_reg action=user_reg.asp method=post>"&_
        		"<input type=hidden name=action value=save>"&_
        		"<input type=hidden name=urlpath value="&urlpath&">"&_

        		"<tr><td>&nbsp;�û���:</td><td><input type=text size=20 name=username>  <input class=button onclick=javascript:window.open('User_RegNameChk.asp?username='+form_reg.username.value,null,'width=60,height=40') href=# type=button value=Check></td></tr>"&_
        		"<tr><td>&nbsp;��  ��:</td><td><input type=password size=20 name=password></td></tr>"&_
        		"<tr><td>&nbsp;�ظ�����:</td><td><input type=password size=20 name=password2></td></tr>"&_
        		"<tr><td>&nbsp;�����ܱ�:</td>"&_
        		"<td>"&_
        		"<select name=question size=1>"&_
        		"		<option value='' selected>--��ѡ��--</option>"&_
        		"		<option value=�ҵĳ������֣�>�ҵĳ������֣�</option>"&_
        		"		<option value=����õ�������˭��>����õ�������˭��</option>"&_
        		"		<option value=����ϲ������ɫ��>����ϲ������ɫ��</option>"&_
        		"		<option value=����ϲ���ĵ�Ӱ��>����ϲ���ĵ�Ӱ��</option>"&_
        		"		<option value=����ϲ����Ӱ�ǣ�>����ϲ����Ӱ�ǣ�</option>"&_
        		"		<option value=����ϲ���ĸ�����>����ϲ���ĸ�����</option>"&_
        		"		<option value=����ϲ����ʳ�>����ϲ����ʳ�</option>"&_
       		 	"		<option value=�����İ��ã�>�����İ��ã�</option>"&_
        		"		<option value=����ѧУ��ȫ����ʲô��>����ѧУ��ȫ����ʲô��</option>"&_
        		"		<option value=�ҵ��������ǣ�>�ҵ��������ǣ�</option>"&_
        		"		<option value=����ϲ����С˵�����֣�>����ϲ����С˵�����֣�</option>"&_
        		"		<option value=����ϲ���Ŀ�ͨ�������֣�>����ϲ���Ŀ�ͨ�������֣�</option>"&_
        		"		<option value=��ĸ�׸��׵����գ�>��ĸ�׸��׵����գ�</option>"&_
        		"		<option value=�������͵�һλ���˵����֣�>�������͵�һλ���˵����֣�</option>"&_
        		"		<option value=����ϲ�����˶���ȫ�ƣ�>����ϲ�����˶���ȫ�ƣ�</option>"&_
        		"		<option value=����ϲ����һ��Ӱ��̨�ʣ�>����ϲ����һ��Ӱ��̨�ʣ�</option>"&_
        		"</select>"&_
        		"</td></tr>"&_
        		"<tr><td>&nbsp;�����:</td><td><input type=text size=20 name=answer></td></tr>"&_
       		 	"<tr><td>&nbsp;����:</td><td><input type=text size=20 name=realname></td></tr>"&_
        		"<tr><td>&nbsp;�Ա�:</td><td><input type=radio value=0 name=sex checked>����&nbsp; &nbsp; <input type=radio value=1 name=sex>&nbsp; Ůʿ</td></tr>"&_
        		"<tr><td>&nbsp;Email:</td><td><input type=text size=20 name=email></td></tr>"&_
        		"<tr><td></td><td><input class=button type=submit value='  �ύ  '></td></tr>"&_
        		"</form>"
call down()
%>
</center>