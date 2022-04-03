<!--#include file="admin_check.asp"-->
<%dim dbpath
dbpath="../"
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_OnOffAdvLeft from root_option where id=1"
rs.open sql,conn,1,1
root_option_OnOffAdvLeft=rs(0)
rs.close
set rs=nothing

action=my_request("action",0)
select case action
case "save"
    call save()
case "add"
    call add()
case "modi"
    call modi()
case "del"
    call del()
end select

sub save()
    root_option_OnOffAdvLeft=my_request("root_option_OnOffAdvLeft",1)
    if root_option_OnOffAdvLeft="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from root_option where id=1"
        rs.open sql,conn,1,3
        rs("root_option_OnOffAdvLeft")=root_option_OnOffAdvLeft
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功设置左侧图标广告开关！","adv_left_set.asp")
    end if
end sub

sub add()
    adv_left_pic   =my_request("adv_left_pic",0)
    adv_left_PicUrl=my_request("adv_left_PicUrl",0)
    if adv_left_pic="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from adv_left"
        rs.open sql,conn,1,3
        rs.addnew
        rs("adv_left_pic")   =adv_left_pic
        rs("adv_left_PicUrl")=adv_left_PicUrl
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功添加一条左侧广告信息！","adv_left_set.asp")
    end if
end sub

sub modi()
   id=my_request("nowid",1)
   adv_left_PicUrl=my_request("adv_left_PicUrl",0)
   if id="" then
       call error()
   else
       sql="update adv_left set adv_left_PicUrl='"&adv_left_PicUrl&"' where adv_left_id="&id
       conn.execute(sql)
       call ok("您已成功修改一条左侧广告信息！","adv_left_set.asp")
  end if
end sub

sub del()
  id=my_request("id",1)
  sql = "delete from adv_left where adv_left_id="&id
  conn.execute(sql)
  call ok("您已成功删除一条左侧广告信息！","adv_left_set.asp")
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>左侧广告-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
<script language="javascript">
//添加验证
function check()
{
 if (document.form1.adv_left_pic.value=="")
	{
	  alert("广告图片不能为空！")
	  document.form1.adv_left_pic.focus()
	  return false
	 }
}

function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="none";
   linkimg2.style.display="none";
  }
  else
  {
   linkimg.style.display="";
   linkimg2.style.display="";
  }
}
</script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form11" action="adv_left_set.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td class="header" colspan="2">左下侧图标广告-控制开关</td>
	</tr>
	<tr>
		<td>左下侧图标广告开关：</td>
		<td>
		    <input type="radio" value="0" name="root_option_OnOffAdvLeft" <%if root_option_OnOffAdvLeft=0 then response.write "checked" %> onClick='showlist("b");'>开启&nbsp;&nbsp;&nbsp; &nbsp; 
		    <input type="radio" value="1" name="root_option_OnOffAdvLeft" <%if root_option_OnOffAdvLeft=1 then response.write "checked" %> onClick='showlist("a");'>关闭&nbsp;&nbsp;&nbsp; &nbsp; 
	
			<font color="#999999">( 选择好后,务必按&quot;提交&quot;按钮才能生效 )</font><tr>
		<td>　</td>
		<td><input type="submit" value="提交" name="B1"></td>
	</tr>
</form>
</tbody>
</table>
<br>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder" id="linkimg" <%if root_option_OnOffAdvLeft=1 then%>style='display:none'<%end if%>>
<tbody class="altbg2">
<form action="adv_left_set.asp" method="post" name="form1" onSubmit="return check();">
<input type="hidden" name="action" value="add">
	<tr>
		<td colspan="2" class="header">左下侧图标广告<b>-添加</b></td>
	</tr>
	<tr>
		<td width="100%" colspan="2" class="altbg1">上传图片尺寸要求:&nbsp;&nbsp; 
		长:140像素&nbsp; 高:任意</td>
	</tr>
	<tr>
		<td width="15%">广告图片上传：</td>
		<td width="85%">
		        <input type="text" name="adv_left_pic" size="30"><input type="button" value="&gt;&gt;点此上传图片" name="action1" onClick="javascript:openWin('Njj_Pic_Upload.asp?Fname=adv_left_pic','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')">
		        &nbsp;&nbsp;尺寸：宽/180 X 高/150</td>
	</tr>
	<tr>
		<td width="15%">广告图片链接：</td>
		<td width="85%">
		<input type="text" name="adv_left_PicUrl" size="50"></td>
	</tr>
	<tr>
		<td width="15%"></td>
		<td width="85%"><input type="submit" value="保存设置" name="B4"></td>
	</tr>
</form>
</tbody>
</table>
<br>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder" id="linkimg2" <%if root_option_OnOffAdvLeft=1 then%>style='display:none'<%end if%>>
<tbody class="altbg2">
	<tr>
		<td colspan="4" height="20" class="header">
		左下侧图标广告<b>-管理</b></td>
	</tr>
	<tr align=center class="altbg1">
		<td class="altbg1">广告图片</td>
		<td class="altbg1">链接</td>
		<td class="altbg1">修改</td>
		<td class="altbg1">删除</td>
	</tr>
    <%
    set rs=server.createobject("adodb.recordset")
    sql="select adv_left_id,adv_left_pic,adv_left_PicUrl from adv_left order by adv_left_id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=4 align=center><font color=red>目前暂无左侧广告条信息,请添加!</font></td></tr>"
    else
        set adv_left_id=rs(0)
        set adv_left_pic=rs(1)
        set adv_left_PicUrl=rs(2)
        while not rs.eof
    %>	
    <form action="adv_left_set.asp" method="post" name="form3">
	<input type="hidden" name="action" value="modi">
    <input type="hidden" name="nowid" value="<%=adv_left_id%>">
	<tr align=center onMouseOver="this.style.backgroundColor='#FFDEAD'" onMouseOut="this.style.backgroundColor=''">
		<td>
		<a href=<%=adv_left_PicUrl%> target=_blank><img src=../uploadpic/<%=adv_left_pic%> border=0></a></td>
		<td><input name=adv_left_PicUrl type=text value="<%=adv_left_PicUrl%>"></td>
		<td><input type="submit" value="修改" name="B6"></td>
		<td><input type="button" onClick="javascript:location.href='adv_left_set.asp?id=<%=adv_left_id%>&action=del';" value="删除" name="B5"></td>
	</tr>
    </form>
    <%
      rs.movenext
      wend
    end if
    rs.close
    set rs=nothing
    %>
</tbody>
</table>

</body>

</html>

