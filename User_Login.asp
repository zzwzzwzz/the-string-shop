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
    window.alert("用户名不能为空！");        
    document.form_1.LoginName.focus();        
    return (false);}  
  
        var filter=/^\s*[@.A-Za-z0-9_-]{3,30}\s*$/;
        if (!filter.test(document.form_1.LoginName.value)) { 
                window.alert("用户名填写不正确,请重新填写！可使用的字符为（A-Z a-z 0-9 _ - .)长度不小于3个字符，不超过30个字符，注意不要使用空格。"); 
                document.form_1.LoginName.focus();
                document.form_1.LoginName.select();
                return (false); 
                }
 if (document.form_1.LoginPass.value == "")        
  {        
    window.alert("密码不能为空！");        
    document.form_1.LoginPass.focus();        
    return (false);}  
  
        var filter=/^\s*[.A-Za-z0-9_-]{5,15}\s*$/;
        if (!filter.test(document.form_1.LoginPass.value)) { 
                window.alert("密码填写不正确,请重新填写！可使用的字符为（A-Z a-z 0-9 _ - .)长度不小于5个字符，不超过15个字符，注意不要使用空格。"); 
                document.form_1.LoginPass.focus();
                document.form_1.LoginPass.select();
                return (false); 
                }
 }

</script>
<%
urlpath=my_request("urlpath",0)

call up("会员登陆","会员登陆","会员登陆")

response.write  "<form name=form_1 action=User_loginCheck.asp method=post onsubmit='return chsubmit();'>"&_
				"<input type=hidden name=urlpath value="&urlpath&">"&_
				"	<tr><td colspan=2 align=center height=40><b>请填写用户名和密码：</b></td></tr>"&_
				"	<tr><td align=right width=40% >&nbsp;用户名:</td><td><input type=text size=14 name=LoginName></td></tr>"&_
				"	<tr><td align=right width=40% >&nbsp;密　码:</td><td><input type=password size=14 name=LoginPass></td></tr>"&_
				"	<tr><td width=40% ></td><td><input class=button type=submit value=' 登 录 '>    <input class=button type=button value=' 注 册 'onclick=window.location='User_Reg.asp?urlpath="&urlpath&"'> <a href=User_PassWordGet.asp>忘记密码</a></td></tr>"&_
				"</form>"
call down()
%></center>