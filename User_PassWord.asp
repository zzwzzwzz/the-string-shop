<center><!--#include file="User_Chk.asp"-->
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file="include/md5.asp"-->
<!--#include file=Sub.asp -->
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

	function check_form()
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
<%
//取出数据
id=session("user_info_id")

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select user_info_RealName,user_info_email,user_info_mobile,user_info_tel,user_info_qq,user_info_msn,user_info_address,user_info_zip from user_info where user_info_id="&id
rs.open sql,conn,1,1
user_info_RealName=rs(0)
user_info_email=rs(1)
user_info_mobile=rs(2)
user_info_tel=rs(3)
user_info_qq=rs(4)
user_info_msn=rs(5)
user_info_address=rs(6)
user_info_zip=rs(7)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call User_PassWordModiSave()
end if

call up("修改密码信息","修改密码信息","修改密码信息")
%>
<!--#include file="User_Menu.asp"-->
<%
response.write  "<form name=form1 action=user_Password.asp method=post>"&_
				"<input type=hidden name=action value=save>"&_
				"<tr><td>旧密码:</td><td><input type=password name=PassWordOld size=20>(必须为 5-10 个字符,只允许字母和数字!)</td></tr>"&_
				"<tr><td>新密码:</td><td><input type=password name=PassWord size=20>(必须为 5-10 个字符,只允许字母和数字!)</td></tr>"&_
				"<tr><td>重复新密码:</td><td><input type=password name=ConfirmPassWord size=20></td></tr>"&_
				"<tr><td></td><td><input type=submit value=提交修改></td></tr>"&_
				"</form>"
call down()
%></center>