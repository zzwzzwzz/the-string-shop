<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=0
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->

<%
action=my_request("action",0)
select case action
    case ""
        call deliver_list() 
    
    case "deliver_add"
        call deliver_add() 
    
    case "deliver_addsave"
        call deliver_addsave() 
    
    case "deliver_modisave"
        call deliver_modisave()
   
    case "deliver_del"
        call deliver_del()
        
end select

sub deliver_addsave()
    root_deliver_name=my_request("root_deliver_name",0)
    root_deliver_day =my_request("root_deliver_day",0)
    root_deliver_cost=my_request("root_deliver_cost",1)
    if root_deliver_name="" or root_deliver_cost="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from root_deliver"
        rs.open sql,conn,1,3
        rs.addnew
        rs("root_deliver_name")=root_deliver_name
        rs("root_deliver_cost")=root_deliver_cost
        rs("root_deliver_day") =root_deliver_day
        rs.update
        rs.close
        set rs=nothing
    
        call ok("您已成功添加一条送货方式信息！","root_deliver_set.asp")
    end if
end sub

sub deliver_modisave()
    id  =my_request("nowid",1)
    root_deliver_name=my_request("root_deliver_name",0)
    root_deliver_day =my_request("root_deliver_day",0)
    root_deliver_cost=my_request("root_deliver_cost",1)
    if root_deliver_name="" or root_deliver_cost="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from root_deliver where id="&id
        rs.open sql,conn,1,3
        rs("root_deliver_name")=root_deliver_name
        rs("root_deliver_cost")=root_deliver_cost
        rs("root_deliver_day") =root_deliver_day
        rs.update
        rs.close
        set rs=nothing
    
        call ok("您已成功修改一条送货方式信息！","root_deliver_set.asp")
    end if
end sub

sub deliver_del()
  id=my_request("id",1)
  sql = "delete from root_deliver where id="&id
  conn.execute(sql)
  call ok("您已成功删除一条送货方式信息！","root_deliver_set.asp")
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>送货方式-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language="javascript">
function check()
{
 if (document.form2.root_deliver_name.value=="")
	{
	  alert("配送方式不能为空！")
	  document.form2.root_deliver_name.focus()
	  return false
	 }
 if (document.form2.root_deliver_cost.value=="")
	{
	  alert("配送费用不能为空！")
	  document.form2.root_deliver_cost.focus()
	  return false
	 }
}
</script>

</head>

<body>

<%sub deliver_list%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="5" class="header">送货方式-设置</td>
	</tr>
    <tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<tr>
		<td class="altbg1">配送方式</td>
		<td class="altbg1">配送费用</td>
		<td class="altbg1">收货时间</td>
		<td class="altbg1">修改保存</td>
		<td class="altbg1">删除</td>
	</tr>
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select id,root_deliver_name,root_deliver_cost,root_deliver_day from root_deliver order by id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
      response.write "<tr><td colspan=5 align=center><font color=red>目前暂无送货方式信息,请<a href=?action=deliver_add>点此添加!</a></font></td></tr>"
    else
      
      set id=rs(0)
      set root_deliver_name=rs(1)
      set root_deliver_cost=rs(2)
      set root_deliver_day=rs(3)
      while not rs.eof
    %>	
    <form action="root_deliver_set.asp" method=post name=form1>
	<input type="hidden" name="action" value="deliver_modisave">
    <input type="hidden" name="nowid" value="<%=id%>">
	<tr>
		<td>
		<input type="text" name="root_deliver_name" size="30" value="<%=root_deliver_name%>"></td>
		<td>
		<input type="text" name="root_deliver_cost" size="8" value="<%=root_deliver_cost%>">
		<font color="#808080">元</font></td>
		<td>
		<input type="text" name="root_deliver_day" size="8" value="<%=root_deliver_day%>">
		<font color="#808080">天</font></td>
		<td>
		<input type="submit" value="修改保存" name="B6"></td>
		<td><input type="button" onclick="javascript:location.href='root_deliver_set.asp?id=<%=id%>&action=deliver_del';" value="删除" name="B5"></td>
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
<br><input type="button" value="送货方式-添加" name="action1" onclick="window.location='root_deliver_set.asp?action=deliver_add'"></p>

<%end sub

sub deliver_add()
%>
<p>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="root_deliver_set.asp" method="post" name="form2" onsubmit="return check();">
<input type="hidden" name="action" value="deliver_addsave">
	<tr>
		<td colspan="2" class="header">送货方式-添加</td>
	</tr>
	<tr>
		<td>配送方式：</td>
		<td><input type="text" name="root_deliver_name" size="30"></td>
	</tr>
	<tr>
		<td>配送费用：</td>
		<td><input type="text" name="root_deliver_cost" size="10">
		<font color="#808080">元</font></td>
	</tr>
	<tr>
		<td>收货时间：</td>
		<td><input type="text" name="root_deliver_day" size="10">
		<font color="#808080">天</font></td>
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

