<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=4
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""��������!"");location.href=""News_Info_List.asp"";</script>")
  response.end
end if

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select news_info_title,news_info_content from news_info where id="&id
rs.open sql,conn,1,1
news_info_title	  = rs(0)
news_info_content = rs(1)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    id				= my_request("id",1)
    news_info_title = my_request("news_info_title",0)
    news_info_content = my_request("Content",0)
    ErrMsg=""
    if id="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>����������ϢID����Ϊ�գ�</li>"
    end if
    if news_info_title="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>������Ϣ���ⲻ��Ϊ�գ�</li>"
    end if
    if news_info_content="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>����������Ϣ���ݲ���Ϊ�գ�</li>"
    end if
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from news_info where id="&id
        rs.open sql,conn,1,3
        rs("news_info_title")   = news_info_title
        rs("news_info_content") = news_info_content
        rs.update
        rs.close
        set rs=nothing
        call ok("���ѳɹ��༭��һ��������Ϣ��","news_info_List.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ϣ-�༭</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language="javascript">
<!--
function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="none";
   linkimg2.style.display="";
  }
  else
  {
   linkimg.style.display="";
   linkimg2.style.display="none";
  }
}

function checkdata()
{
if (document.form1.viewhtml.checked == true)
	{
	  alert("�Բ�����ȡ�����鿴HTMLԴ���롱�������ӣ�")
	  document.form1.viewhtml.focus()
	  return false
	 }
}
//-->
</script>

</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="News_Info_Modi.asp" method="post" name="form1" onsubmit="return checkdata();">
<input type="hidden" name="action" value="save">
<input type="hidden" name="id" value="<%=id%>"> 
	<tr>
		<td colspan="2" class="header">������Ϣ-�༭</td>
	</tr>
	<tr>
		<td>���⣺</td>
		<td>
		<input type="text" name="news_info_title" size="40" value="<%=news_info_title%>"></td>
	</tr>
	<tr>
		<td height="23">�������ݣ�</td>
		<td height="23">
		    <textarea cols=80 rows=20 id="content" name="Content"><%= Server.HTMLEncode(news_info_content) %></textarea>
        </td>
	</tr>
	<tr>
		<td>��</td>
		<td>
		    <input type="submit" value="  ��  ��  " name="B1">&nbsp; 
		    <input type="reset" value="  ��  ��  " name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>