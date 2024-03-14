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
function noChar(element1){//���зǷ��ַ� ���� true
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
			alert("����д�����룡");
			frm.PassWordOld.focus();
			return false;			
		}
		
        if(frm.PassWord.value=="") 
		{
			alert("����д�����룡");
			frm.PassWord.focus();
			return false;			
		}
		
		if(frm.PassWordOld.value.length<5 || frm.PassWordOld.value.length>10) 
		{
			alert("���������Ϊ 5-10 ���ַ���ֻ������ĸ�����֣���");
			frm.PassWord.focus();
			return false;			
		}

		if(frm.PassWord.value.length<5 || frm.PassWord.value.length>10) 
		{
			alert("���������Ϊ 5-10 ���ַ���ֻ������ĸ�����֣���");
			frm.PassWord.focus();
			return false;			
		}
		
		if(frm.PassWord.value!=frm.ConfirmPassWord.value) 
		{
			alert("ȷ����������������һ�£�");
			frm.ConfirmPassWord.focus();
			return false;			
		}
		frm.Submit1.value = "�ύ�У����Ժ�..." 
	    frm.Submit1.disabled = true;	
		frm.submit();		
	}	
	
</script>
<%
'ȡ������
id=session("user_info_id")

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select user_info_RealName,user_info_email,user_info_mobile,user_info_address,user_info_zip from user_info where user_info_id="&id
rs.open sql,conn,1,1
user_info_RealName=rs(0)
user_info_email=rs(1)
user_info_mobile=rs(2)
user_info_address=rs(3)
user_info_zip=rs(4)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call User_PassWordModiSave()
end if

call up("�޸�������Ϣ","�޸�������Ϣ","�޸�������Ϣ")
%>
<!--#include file="User_Menu.asp"-->
<%
response.write  "<form name=form1 action=user_Password.asp method=post>"&_
				"<input type=hidden name=action value=save>"&_
				"<tr><td>������:</td><td><input type=password name=PassWordOld size=20>(����Ϊ 5-10 ���ַ���ֻ������ĸ������)</td></tr>"&_
				"<tr><td>������:</td><td><input type=password name=PassWord size=20>(����Ϊ 5-10 ���ַ���ֻ������ĸ������)</td></tr>"&_
				"<tr><td>�ظ�������:</td><td><input type=password name=ConfirmPassWord size=20></td></tr>"&_
				"<tr><td></td><td><input type=submit value=�ύ�޸�></td></tr>"&_
				"</form>"
call down()
%></center>