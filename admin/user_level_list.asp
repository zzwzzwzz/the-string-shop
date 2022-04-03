<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=3
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->

<%
action=my_request("action",0)
select case action
    case ""
        call level_list() 
    
    case "level_add"
        call level_add() 
    
    case "level_addsave"
        call level_addsave() 
    
    case "level_modisave"
        call level_modisave()
   
    case "level_del"
        call level_del()
        
end select

sub level_addsave()
    user_level_name=my_request("user_level_name",0)
    user_level_markMin=my_request("user_level_markMin",0)
    user_level_markMax=my_request("user_level_markMax",0)
    user_level_rebate=my_request("user_level_rebate",0)
    if user_level_name="" or user_level_markMin="" or user_level_markMax="" or user_level_rebate="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from user_level"
        rs.open sql,conn,1,3
        rs.addnew
        rs("user_level_name")=user_level_name
        rs("user_level_markmin")=user_level_markmin
        if user_level_markmax<>"" then
        	rs("user_level_markmax")=user_level_markmax
        end if
        rs("user_level_rebate")=user_level_rebate
        rs.update
        rs.close
        set rs=nothing
    
        call ok("您已成功添加一条会员级别信息！","user_level_list.asp")
    end if
end sub

sub level_modisave()
    id  =my_request("nowid",1)
    user_level_name=my_request("user_level_name",0)
    user_level_markMin=my_request("user_level_markMin",0)
    user_level_markMax=my_request("user_level_markMax",1)
    user_level_rebate=my_request("user_level_rebate",1)
    if id="" or user_level_name="" or user_level_markMin="" or user_level_markMax="" or user_level_rebate="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from user_level where id="&id
        rs.open sql,conn,1,3
        rs("user_level_name")=user_level_name
        rs("user_level_markmin")=user_level_markmin
        rs("user_level_markmax")=user_level_markmax
        rs("user_level_rebate")=user_level_rebate
        rs.update
        rs.close
        set rs=nothing
    
        call ok("您已成功修改一条会员级别信息！","user_level_list.asp")
    end if
end sub

sub level_del()
  id=my_request("id",1)
  sql = "delete from user_level where id="&id
  conn.execute(sql)
  call ok("您已成功删除一条会员级别信息！","user_level_list.asp")
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>送货方式-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
<script language="javascript">
function check()
{
 if (document.form2.user_level_name.value=="")
	{
	  alert("级别名称不能为空！")
	  document.form2.user_level_name.focus()
	  return false
	 }
 if (document.form2.user_level_rebate.value=="")
	{
	  alert("折扣不能为空！")
	  document.form2.user_level_rebate.focus()
	  return false
	 }
}
</script>

</head>

<body>

<%sub level_list%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="6" class="header">会员级别-设置</td>
	</tr>
	<tr>
		<td class="altbg1">会员级别名称</td>
		<td class="altbg1">积分下限</td>
		<td class="altbg1">积分上限<span style="font-weight: 400">(99999999表示不设上限)</span></td>
		<td class="altbg1">享受购买总价(不含配送费)折扣</td>
		<td class="altbg1">修改保存</td>
		<td class="altbg1">删除</td>
	</tr>
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select id,user_level_name,user_level_markmin,user_level_markmax,user_level_rebate from user_level order by id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
      response.write "<tr><td colspan=6 align=center><font color=red>目前暂无会员级别信息,请<a href=?action=level_add>点此添加!</a></font></td></tr>"
    else
      set id=rs(0)
      set user_level_name=rs(1)
      set user_level_markmin=rs(2)
      set user_level_markmax=rs(3)
      set user_level_rebate=rs(4)
      while not rs.eof
    %>	
    <form action="user_level_list.asp" method=post name=form1>
	<input type="hidden" name="action" value="level_modisave">
    <input type="hidden" name="nowid" value="<%=id%>">
	<tr>
		<td>
		<input type="text" name="user_level_name" size="16" value="<%=user_level_name%>"></td>
		<td>
		<input type="text" name="user_level_markmin" size="8" value="<%=user_level_markmin%>">
		<font color="#808080">分</font></td>
		<td>
		<input type="text" name="user_level_markmax" size="8" value="<%=user_level_markmax%>">
		<font color="#808080">分</font></td>
		<td>
		<input type="text" name="user_level_rebate" size="4" value="<%=user_level_rebate%>">
		折</td>
		<td>
		<input type="submit" value="修改保存" name="B6"></td>
		<td><input type="button" onclick="javascript:location.href='user_level_list.asp?id=<%=id%>&action=level_del';" value="删除" name="B5"></td>
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
<br>
<input type="button" value="会员级别-添加" name="action1" onclick="window.location='user_level_list.asp?action=level_add'"></p>

<%end sub

sub level_add()
%>
<p>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="user_level_list.asp" method="post" name="form2" onsubmit="return check();">
<input type="hidden" name="action" value="level_addsave">
	<tr>
		<td colspan="2" class="header">会员级别-添加</td>
	</tr>
	<tr>
		<td>会员级别名称：</td>
		<td><input type="text" name="user_level_name" size="30"></td>
	</tr>
	<tr>
		<td>此级别要求积分范围：</td>
		<td>积分下限:<input type="text" name="user_level_MarkMin" size="10">&nbsp; 至&nbsp; 
		积分上限:<input type="text" name="user_level_MarkMax" size="10">&nbsp;&nbsp;&nbsp; 
		<br>
		<%
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select root_option_MarkYuan from root_option where id=1"
		rs.open sql,conn,1,1
		root_option_MarkYuan       = rs(0)
		rs.close
		set rs=nothing
		%><font color="#808080">没有上限,请填写</font><span style="font-weight: 400"><font color="#FF6600">99999999
		</font>(共8个9)</span><font color="#808080"><br>
		已设定购物1元=<%=root_option_MarkYuan%>积分,<a href="user_Option_Set.asp"><font color="#0000FF">点此进行积分换算设置</font></a></font></td>
	</tr>
	<tr>
		<td>享受购买总价（不含配送费）折扣：</td>
		<td><input type="text" name="user_level_rebate" size="5">折&nbsp;&nbsp; 
		(此格式要求：例：<font color="#FF6600">95</font> 折;注意:100=不打折)</td>
	</tr>
	<tr>
		<td>　</td>
		<td>　</td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="提交" name="B1">&nbsp;
		<input type="reset" value="重置" name="B2"></td>
	</tr>
</form>
</tbody>
</table>
<%end sub%>
</body>

</html>

