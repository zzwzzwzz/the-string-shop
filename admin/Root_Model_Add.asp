<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=8
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
if action="save" then
   call save()
end if

sub save()
    root_model_name=my_request("root_model_name",0)
    root_model_css=my_request("root_model_css",0)
    ErrMsg=""
    if root_model_name="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>ģ�����Ʋ���Ϊ�գ�</li>"
    end if
    if root_model_css="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��ʽ���ļ�������Ϊ�գ�</li>"
    end if
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from root_model"
        rs.open sql,conn,1,3
        rs.addnew
        rs("root_model_name")=root_model_name
        rs("root_model_css")=root_model_css
        rs.update
        rs.close
        set rs=nothing
        call ok("���ѳɹ������һ����վģ����Ϣ��","root_model_list.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ģ��-���</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language = "JavaScript">  
function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="none";
  }
  else
  {
   linkimg.style.display="";
  }
}
</script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="Root_Model_Add.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">ģ��-���</td>
	</tr>
	<tr>
		<td>ģ�����ƣ�</td>
		<td>
		    <input type="text" name="root_model_name" size="20"></td>
	</tr>
	<tr>
		<td>ģ����ʽ��-�ļ�����</td>
		<td>
		    <input type="text" name="root_model_css" size="20">.css<font color="#808080">&nbsp;&nbsp;
			<br>
			��ȷ�����ѽ����ļ��ŵ���styleĿ¼��;<br>
			��ģ���õ���ͼƬ�ļ���Ҳ��һ���ŵ�styleĿ¼��;</font></td>
	</tr>
	<tr>
		<td>��</td>
		<td>
		   <input type="submit" value="  ��  ��  " name="Submit1">&nbsp; 
		   <input type="reset" value="  ��  ��  " name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

