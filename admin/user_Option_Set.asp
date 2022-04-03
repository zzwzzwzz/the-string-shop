<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=3
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_MarkYuan,root_option_PriceShowType from root_option where id=1"
rs.open sql,conn,1,1
root_option_MarkYuan       = rs(0)
root_option_PriceShowType  = rs(1)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
   call save()
end if

sub save()
    root_option_MarkYuan   = my_request("root_option_MarkYuan",1)
    root_option_PriceShowType= my_request("root_option_PriceShowType",1)
    
    ErrMsg=""
    if root_option_MarkYuan="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>商品缩图尺寸-横宽不能为空！</li>"
    end if
    if root_option_PriceShowType="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>商品缩图尺寸-竖高不能为空！</li>"
    end if
    
    if FoundErr<>True then
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from root_option where id=1"
        rs.open sql,conn,1,3
        rs("root_option_MarkYuan")     	= root_option_MarkYuan
        rs("root_option_PriceShowType") = root_option_PriceShowType
        rs.update
        rs.close
        set rs=nothing

        call ok("您已成功保存会员选项设置！","user_option_set.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员选项-设置</title>
<link rel="stylesheet"  href="style.css" type="text/css">
<script language="JavaScript" type="text/JavaScript">
function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="";
  }
  else
  {
   linkimg.style.display="none";
  }
}
</script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="user_Option_Set.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">会员选项-设置</td>
	</tr>
	<tr>
		<td>当购物满<font color="#FF6600"><b>1</b></font>元时，可获得积分数：</td>
		<td>
		<input type="text" name="root_option_MarkYuan" size="3" value="<%=root_option_MarkYuan%>">分 </td>
	</tr>
	<tr>
		<td>前台商品介绍页-商品不同会员价格显示方式：</td>
		<td>
			<input type="radio" value="0" name="root_option_PriceShowType" <%if root_option_PriceShowType=0 then response.write "checked"%>>只显示网站基准价格<font color="#808080"> 
			(即商品发布页中的网站价)</font><br>
		<input type="radio" value="1" name="root_option_PriceShowType" <%if root_option_PriceShowType=1 then response.write "checked"%>>会员登陆后显示同级别及以下级别会员价格<br>
		<input type="radio" value="2" name="root_option_PriceShowType" <%if root_option_PriceShowType=2 then response.write "checked"%>>显示所有级别会员价格</td>
	</tr>
	</td></tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="  提  交  " name="B1"></td>
	</tr>
</form>
</tbody>
</table>


</body>

</html>
