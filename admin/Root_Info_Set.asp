<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=0
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select top 1 root_info_OnOff,root_info_OffNote,root_info_LogoPic,root_info_tel,root_info_email,root_info_IndexKeyWords,root_info_IndexTitle,root_info_IndexDescription,root_info_sitename,root_info_address,root_info_zip from root_info"
rs.open sql,conn,1,1
root_info_OnOff   =rs(0)
root_info_OffNote =rs(1)
root_info_LogoPic =rs(2)
root_info_tel     =rs(3)
root_info_email   =rs(4)
root_info_IndexKeyWords=rs(5)
root_info_IndexTitle=rs(6)
root_info_IndexDescription=rs(7)
root_info_sitename =rs(8)
root_info_address =rs(9)
root_info_zip=rs(10)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    root_info_OnOff           =my_request("root_info_OnOff",1)
    root_info_OffNote         =my_request("root_info_OffNote",0)
    root_info_LogoPic         =my_request("root_info_LogoPic",0)
    root_info_tel             =my_request("root_info_tel",0)
    root_info_email           =my_request("root_info_email",0)
    root_info_IndexTitle      =my_request("root_info_IndexTitle",0)
    root_info_IndexKeyWords   =my_request("root_info_IndexKeyWords",0)
    root_info_IndexDescription=my_request("root_info_IndexDescription",0)
    root_info_sitename			=my_request("root_info_sitename",0)
    root_info_address			=my_request("root_info_address",0)
    root_info_zip				=my_request("root_info_zip",0)
           
    if root_info_LogoPic="" then
        response.redirect "error.htm"
        response.end
    else
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from root_info where id=1"
        rs.open sql,conn,1,3
        rs("root_info_OnOff")           =root_info_OnOff
        rs("root_info_OffNote")         =root_info_OffNote
        rs("root_info_LogoPic")         =root_info_LogoPic
        rs("root_info_tel")             =root_info_tel
        rs("root_info_email")           =root_info_email
        rs("root_info_IndexTitle")      =root_info_IndexTitle
        rs("root_info_IndexKeyWords")   =root_info_IndexKeyWords
        rs("root_info_IndexDescription")=root_info_IndexDescription
        rs("root_info_sitename")=root_info_sitename
        rs("root_info_address")=root_info_address
        rs("root_info_zip")=root_info_zip
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功保存基本资料设置！","root_info_set.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>基本-基本资料-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language="JavaScript" type="text/JavaScript">
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

function check_form()
	{
		var frm;
		frm=document.form1;
		if(frm.root_info_SiteName.value=="") 
		{
			alert("请填写网店名称！");
			frm.root_info_SiteName.focus();
			return false;			
		}
		frm.Submit1.value = "提交中，请稍候..." 
	    frm.Submit1.disabled = true;	
		frm.submit();		
}	
</script>
<style>
<!--
.i_table {BORDER: #3191bc 1px solid;background:#afd7e1;}
.tpc_title { font-size: 12px;font-weight:bold;}
-->
</style>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="root_info_set.asp" method="post" onsubmit="return check_form();">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td class="header" colspan="2">基本资料-设置</td>
	</tr>
	<tr>
		<td>网站开关：</td>
		<td>
		    <input type="radio" value="0" name="root_info_OnOff" <%if root_info_OnOff=0 then response.write "checked" %> onClick='showlist("a");'>开启&nbsp;&nbsp; 
		    <input type="radio" value="1" name="root_info_OnOff" <%if root_info_OnOff=1 then response.write "checked" %> onClick='showlist("b");'>关闭</td>
	</tr>
	
	<tr id="linkimg" <%if root_info_OnOff=0 then%>style='display:none'<%end if%>>
        <td>网站关闭时的显示语：</td>
		<td><textarea rows="4" name="root_info_OffNote" cols="47"><%=root_info_OffNote%></textarea></td>
	</tr>
	<tr id="linkimg2" <%if root_info_OnOff=1 then%>style='display:none'<%end if%>>
	  <td colspan=2>
	    <table cellpadding="4" style="border-collapse: collapse" border="1" bordercolor="#CCCCCC" width="100%">
	      <tr>
		      <td>网站名称：</td>
		      <td>
				<input type="text" name="root_info_SiteName" size="30" value="<%=root_info_SiteName%>"></td>
	      </tr>
	      <%if root_info_LogoPic<>"" then%>
	      <tr>
		      <td>网站LOGO：</td>
		      <td><img src=../uploadpic/<%=root_info_LogoPic%> border=0></td>
	      </tr>
	      <%end if%>
	      <tr>
		      <td>LOGO上传：</td>
		      <td><input type="text" name="root_info_LogoPic" size="30" value="<%=root_info_LogoPic%>">
		          <input type="button" value="点此上传" name="action0" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=root_info_LogoPic','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=400,height=100')">
		      </td>
	      </tr>
	      <tr>
		      <td>联系地址：</td>
		      <td>
				<input type="text" name="root_info_address" size="30" value="<%=root_info_address%>"></td>
	      </tr>
	      <tr>
		      <td>邮政编码：</td>
		      <td>
				<input type="text" name="root_info_zip" size="30" value="<%=root_info_zip%>"></td>
	      </tr>
	      <tr>
		      <td>联系电话：</td>
		      <td><input type="text" name="root_info_tel" size="30" value="<%=root_info_tel%>"></td>
	      </tr>
	      <tr>
		      <td>E-mail：</td>
		      <td><input type="text" name="root_info_email" size="30" value="<%=root_info_email%>"></td>
	      </tr>
	      <tr>
		      <td colspan="2" bgcolor="#654321" ><b><font color="#FFFFFF">网站首页-优化 (有利于搜索引擎收录及排名靠前)</font></b></td>
	     </tr>
	      <tr>
		      <td>首页标题：<font color="#808080">(不超过20个汉字)</font><br>
				<font color="#808080">2-3个主营商品的关键词<br>
		      <td>
		      <input type="text" name="root_info_IndexTitle" size="30" value="<%=root_info_IndexTitle%>"></td>
	     </tr>
	      <tr>
		      <td>网站关键字：<font color="#808080">(不超过20个汉字)</font></td>
		      <td>
		      <input type="text" name="root_info_IndexKeyWords" size="30" value="<%=root_info_IndexKeyWords%>"></td>
	     </tr>
	      <tr>
		      <td>网站描述：<font color="#808080">(不超过25个汉字)</font></td>
		      <td>
		      <input type="text" name="root_info_IndexDescription" size="30" value="<%=root_info_IndexDescription%>"></td>
	     </tr>
	    	</table>
	  </td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="提交" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="reset" value="重置" name="B2"></td>
	</tr>
	</form>
 </tbody>
</table>
<br>
</body>

</html>

