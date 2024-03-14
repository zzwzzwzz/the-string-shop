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
    window.alert("用户名不能为空！");        
    document.form1.LoginName.focus();        
    return (false);}  
  
        var filter=/^\s*[@.A-Za-z0-9_-]{3,30}\s*$/;
        if (!filter.test(document.form1.LoginName.value)) { 
                window.alert("用户名填写不正确,请重新填写！可使用的字符为（A-Z a-z 0-9 _ - .)长度不小于3个字符，不超过30个字符，注意不要使用空格。"); 
                document.form1.LoginName.focus();
                document.form1.LoginName.select();
                return (false); 
                }
 if (document.form1.LoginPass.value == "")        
  {        
    window.alert("密码不能为空！");        
    document.form1.LoginPass.focus();        
    return (false);}  
  
        var filter=/^\s*[.A-Za-z0-9_-]{5,15}\s*$/;
        if (!filter.test(document.form1.LoginPass.value)) { 
                window.alert("密码填写不正确,请重新填写！可使用的字符为（A-Z a-z 0-9 _ - .)长度不小于5个字符，不超过15个字符，注意不要使用空格。"); 
                document.form1.LoginPass.focus();
                document.form1.LoginPass.select();
                return (false); 
                }
 }
</script>
<%
call up("结算身份选择","结算身份选择","<a href=cart_list.asp>购物车</a> &raquo; 结算身份选择")

response.write  "<tr><td colspan=2 height=8></td></tr>"&_
				"<tr>"&_
				"	<td width=50% valign=top style='border-right: 1px solid #CCCCCC'>"&_
				"		<table width=100% ><form action=User_LoginCheck.asp method=post name=form1 onsubmit=return chsubmit1();>"&_
				"			<input type=hidden name=urlpath value="&urlpath&">"&_
				"			<tr><td colspan=2>&nbsp;&nbsp;<b>以会员身份结算订单：</b></td></tr>"&_
				"			<tr><td>&nbsp;&nbsp;&nbsp;用户名:</td><td><input type=text size=14 name=loginname></td></tr>"&_
				"			<tr><td>&nbsp;&nbsp;&nbsp;密　码:</td><td><input type=password size=14 name=loginpass></td></tr>"&_
				"			<tr><td></td><td><input type=submit value=登录> <input type=button value=注册 onclick=window.location='User_Reg.asp?urlpath="&urlpath&"'> <a href=Member_PassWordGet.asp>忘记密码？</a></td></tr>"&_
				"		</form></table>"&_
				"	</td>"&_
				"	<td width=50% valign=top>"&_
				"		<table width=100% ><tr><td>&nbsp;&nbsp;<b>以游客身份结算订单：</b></td></tr>"&_
				"		<tr><td>&nbsp;&nbsp;<input onclick=document.location.href='Cart_Order.asp'; type=button value=结算></td></tr></table>"&_
				"	</td>"&_
				"</tr>"

call down()
%>
</center>
</center>