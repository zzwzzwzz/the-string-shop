<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=9
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/md5.asp"-->
<%
action=my_request("action",0)
if action="save" then
    call Save()
end if

'管理人员密码-修改保存
sub Save() 
    passwordold=my_request("PassWordOld",0)
	passw=request("passw")
    password=my_request("password",0)
    confirmpassword=my_request("confirmpassword",0)
    if session("admin_info_UserName")="" or passwordold="" or password="" or confirmpassword="" then
        Response.write "<script>alert(""对不起，您的修改信息不完整，请重新填写。"");"
        response.write"javascript:history.go(-1)</SCRIPT>"
        Response.end
    end if
    if password<>confirmpassword then
        response.write"<SCRIPT language=JavaScript>alert('两次新密码输入不一致！');"
        response.write"javascript:history.go(-1)</SCRIPT>"
        response.end
    end if

    Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select * from admin_info where admin_info_UserName='"&session("admin_info_UserName")&"'"
    rs.open sql,conn,1,3

    if passw <> md5(passwordold,32) then
        response.write"旧密码输入有错误！"
        response.write"<span onclick='history.back()'>返回</span>"
        response.end
    Else
    rs("admin_info_PassWord")=md5(password,32)
        rs.update
    end If
   rs.close
    set rs=Nothing
    

    Response.write "<script>alert(""管理密码已成功修改"");location.href=""admin_info_list.asp"";</script>"
    Response.end    
end sub
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理员-管理人员密码-修改</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language="JavaScript" type="text/JavaScript">
function noChar(element1){//含有非法字符 返回 true
   text="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890._-@";
   for(i=0;i<=element1.length-1;i++){
      char1=element1.charAt(i);
      index=text.indexOf(char1);
      if(index==-1)
        return true;
   }
   return false;
}

	function check()
	{

		var frm;
		frm=document.form1;
		if(frm.PassWordOld.value=="") 
		{
			alert("请填写旧密码！");
			frm.PassWordOld.focus();
			return false;			
		}
		
        if(frm.PassWord.value=="") 
		{
			alert("请填写新密码！");
			frm.PassWord.focus();
			return false;			
		}
		
		if(frm.PassWordOld.value.length<5 || frm.PassWordOld.value.length>10) 
		{
			alert("旧密码必须为 5-10 个字符（只允许字母和数字）！");
			frm.PassWord.focus();
			return false;			
		}

		if(frm.PassWord.value.length<5 || frm.PassWord.value.length>10) 
		{
			alert("新密码必须为 5-10 个字符（只允许字母和数字）！");
			frm.PassWord.focus();
			return false;			
		}
		
		if(frm.PassWord.value!=frm.ConfirmPassWord.value) 
		{
			alert("确认密码和新密码必须一致！");
			frm.ConfirmPassWord.focus();
			return false;			
	
		}
		frm.Submit1.value = "提交中，请稍候..." 
	    frm.Submit1.disabled = true;	
		frm.submit();		
	}	
	
</script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="?action=save" method="post" name="form1" onsubmit="return check();">
 <%Set rss= Server.CreateObject("ADODB.Recordset")
    sql="select * from admin_info where admin_info_UserName='"&session("admin_info_UserName")&"'"
    rss.open sql,conn,1,1%><input name="passw" type=hidden value="<%=rss("admin_info_PassWord")%>"><%rss.close
	Set rss=nothing%>
	<tr>
		<td colspan="2" class="header">管理人员密码-修改</td>
	</tr>
	<tr>
		<td>登陆用户名：</td>
		<td><font color="#FF0000"><%=session("admin_info_UserName")%></font>
		<font color="#808080">(注：用户名不可修改)</font></td>
	</tr>
	<tr>
		<td>旧登陆密码：</td>
		<td>
		<input type="password" name="PassWordOld" size="20"></td>
	</tr>
	<tr>
		<td>新登陆密码：</td>
		<td>
		<input type="password" name="PassWord" size="20"></td>
	</tr>
	<tr>
		<td>再输一次新密码：</td>
		<td>
		<input type="password" name="ConfirmPassWord" size="20"></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="提交" name="Submit1">&nbsp;
		<input type="reset" value="重置" name="B2"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
 
