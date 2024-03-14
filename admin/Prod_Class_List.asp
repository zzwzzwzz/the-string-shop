<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->

<%
action=my_request("action",0)
select case action
 case ""
   call classlist()
   
 case "b_add"
   call b_add()
   
 case "s_add"
   call s_add()
   
 case "b_addsave"
   call b_addsave()
   
 case "s_addsave"
   call s_addsave() 
   
 case "b_modi"
   call b_modi()
   
 case "b_modisave"
   call b_modisave()
   
 case "s_modi"
   call s_modi()
   
 case "s_modisave"
   call s_modisave()
   
 case "b_del"
   call b_del()
   
 case "s_del"
   call s_del()
   
end select

sub b_addsave()
    prod_BigClass_name=my_request("prod_BigClass_name",0)
    prod_BigClass_sort=my_request("prod_BigClass_sort",1)
    if prod_BigClass_name="" or prod_BigClass_sort="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from prod_BigClass where prod_BigClass_sort="&prod_BigClass_sort
        rs.open sql,conn,1,3
        if not rs.eof then
            response.write "<SCRIPT language=JavaScript>alert('排序重复，请检查后重新提交。');"
    		response.write "location.href='javascript:history.go(-1)';</SCRIPT>"
    		response.end
        end if
        rs.close
        set rs=nothing
        
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from prod_BigClass"
        rs.open sql,conn,1,3
        rs.addnew
        rs("prod_BigClass_name")=prod_BigClass_name
        rs("prod_BigClass_sort")=prod_BigClass_sort
        rs.update
        rs.close
        set rs=nothing
   		call ok("您已成功添加一个商品大类！","Prod_Class_List.asp")
    end if
end sub

sub s_addsave()
  	prod_SmallClass_bid=my_request("prod_SmallClass_bid",1)
  	prod_SmallClass_name=my_request("prod_SmallClass_name",0)
  	prod_SmallClass_sort=my_request("prod_SmallClass_sort",1)
  	if prod_SmallClass_name="" or prod_SmallClass_bid="" then
    	call error()
 	 else
    	Set rs= Server.CreateObject("ADODB.Recordset")
    	sql="select * from prod_SmallClass"
    	rs.open sql,conn,1,3
    	rs.addnew
    	rs("prod_SmallClass_bid")=prod_SmallClass_bid
    	rs("prod_SmallClass_name")=prod_SmallClass_name
    	rs("prod_SmallClass_sort")=prod_SmallClass_sort
    	rs.update
    	rs.close
    	set rs=nothing
    	call ok("您已成功添加一个商品小类！","Prod_Class_List.asp")
  	end if
end sub

sub b_modisave()
    prod_BigClass_id=my_request("prod_BigClass_id",1)
    prod_BigClass_name=my_request("prod_BigClass_name",0)
    prod_BigClass_sort=my_request("prod_BigClass_sort",1)
    if prod_BigClass_id="" or prod_BigClass_name="" or prod_BigClass_sort="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from prod_BigClass where prod_BigClass_id="&prod_BigClass_id
        rs.open sql,conn,1,3
        rs("prod_BigClass_name")=prod_BigClass_name
        rs("prod_BigClass_sort")=prod_BigClass_sort
        rs.update
        rs.close
        set rs=nothing
   		call ok("您已成功修改一个商品大类！","Prod_Class_List.asp")
    end if
end sub

sub s_modisave()
  prod_SmallClass_id=my_request("prod_SmallClass_id",1)
  prod_SmallClass_name=my_request("prod_SmallClass_name",0)
  prod_SmallClass_sort=my_request("prod_SmallClass_sort",0)
  
  if prod_SmallClass_id="" or prod_SmallClass_name="" or prod_SmallClass_sort="" then
      call error()
  else
      Set rs= Server.CreateObject("ADODB.Recordset")
      sql="select * from prod_SmallClass where prod_SmallClass_id="&prod_SmallClass_id
      rs.open sql,conn,1,3
      rs("prod_SmallClass_name")=prod_SmallClass_name
      rs("prod_SmallClass_sort")=prod_SmallClass_sort
      rs.update
      rs.close
      set rs=nothing
      call ok("您已成功修改一个商品小类！","Prod_Class_List.asp")
  end if
end sub

sub b_del()
   prod_BigClass_id=my_request("prod_BigClass_id",1)
   conn.execute ("delete from [prod_BigClass] where prod_BigClass_id="&prod_BigClass_id)
   conn.execute ("delete from [prod_SmallClass] where prod_smallClass_bid="&prod_BigClass_id)
   conn.execute ("delete from [prod_info] where prod_info_bid="&prod_BigClass_id)
   call ok("您已成功删除一个商品大类！","Prod_Class_List.asp")
end sub

sub s_del()
   prod_SmallClass_id=my_request("prod_SmallClass_id",1)
   conn.execute ("delete from [prod_SmallClass] where prod_SmallClass_id="&prod_SmallClass_id)
   conn.execute ("delete from [prod_info] where prod_info_sid="&prod_SmallClass_id)
   call ok("您已成功删除一个商品小类！","Prod_Class_List.asp")
end sub
%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品类别-管理</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<!-- 底部开始 -->
<%sub classlist()%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="Prod_add.asp" method="post" onsubmit="return check_form();">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td class="header" colspan=2>商品类别- 管理</td>
	</tr>
	<tr>
		<td colspan=2><a href="prod_class_list.asp?action=b_add">添加一级分类</a></td>
	</tr>
    <tr>
    <%
    set rs=server.createobject("adodb.recordset")
    sql="select prod_BigClass_id,prod_BigClass_name,prod_BigClass_sort from prod_BigClass order by prod_BigClass_sort asc"
    rs.open sql,conn,1,1
    if rs.eof then 
      response.write "<tr><td align=center colspan=2><font color=red>目前暂无商品大类信息,请<a href=Prod_Class_List.asp?action=b_add>添加一级分类</a></font></td></tr>"
    else
      i=1
      
      set prod_BigClass_id     =rs(0)
      set prod_BigClass_name   =rs(1)
      set prod_BigClass_sort=rs(2)
      while not rs.eof
    %>
		<td valign="top"><img border="0" src="images/icon_arrow.gif" width="4" height="7">&nbsp;&nbsp;<img border="0" src="images/tree_folder-.gif" width="15" height="15"><b><%=prod_BigClass_name%></b> 
		(<a href="Prod_Class_List.asp?action=s_add&prod_BigClass_id=<%=prod_BigClass_id%>&prod_BigClass_name=<%=prod_BigClass_name%>">添加二级分类</a>)
		<a href="?action=b_modi&prod_BigClass_id=<%=prod_BigClass_id%>&prod_BigClass_name=<%=prod_BigClass_name%>&prod_BigClass_sort=<%=prod_BigClass_sort%>">修改</a>  <a href="?action=b_del&prod_BigClass_id=<%=prod_BigClass_id%>" onclick="{if(confirm('大类删除后将同时删除所有此大类下的商品信息及小类信息,并且无法恢复，您确定要删除选定的大类吗？')){this.document.form1.submit();return true;}return false;}">删除</a> 
		<table border="0" width="100%" cellpadding="0" style="border-collapse: collapse">
		<%
        set rs1=server.createobject("adodb.recordset")
        sql1="select prod_SmallClass_id,prod_SmallClass_name,prod_smallclass_sort from prod_SmallClass where prod_SmallClass_bid="&rs("prod_BigClass_id")&" order by prod_SmallClass_sort asc"
        rs1.open sql1,conn,1,1
        if rs1.eof then 
          response.write "<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>此大类下暂无商品小类信息,请<a href=Prod_Class_List.asp?action=s_add&prod_BigClass_id="&prod_BigClass_id&"&prod_BigClass_name="&prod_BigClass_name&">添加二级分类</a></font></td></tr>"
        else
          
          set prod_SmallClass_id=rs1(0)
          set prod_SmallClass_name=rs1(1)
          set prod_smallclass_sort=rs1(2)
          while not rs1.eof
        %>
            <tr>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="images/tree_line.gif" width="17" height="16"><img border="0" src="images/tree_folder-.gif" width="15" height="15"><%=prod_SmallClass_name%> <a href=?action=s_modi&prod_SmallClass_id=<%=prod_SmallClass_id%>&prod_SmallClass_name=<%=prod_SmallClass_name%>&prod_BigClass_id=<%=prod_BigClass_id%>&prod_BigClass_name=<%=prod_BigClass_name%>&prod_smallclass_sort=<%=prod_smallclass_sort%>>修改</a> <a href="?action=s_del&prod_SmallClass_id=<%=prod_SmallClass_id%>" onclick="{if(confirm('小类删除后将同时删除所有此小类下的商品信息,并且无法恢复，您确定要删除选定的小类吗？')){this.document.form1.submit();return true;}return false;}">删除</a></td>
			</tr>
	    <%
          rs1.movenext
          wend
       end if
       rs1.close
       set rs1=nothing
       %>
	  </table>
	  </td>
    <%
	  if (i mod 2)=0 then
	    response.write "</tr><tr><td colspan=2><hr></td></tr>"
	  end if
	  rs.movenext
	  i=i+1
	  wend
	end if
    rs.close
    set rs=nothing
    %>
</tbody>
</table>
<br>
<%end sub

sub b_add()
%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Prod_Class_List.asp" method="post" name="form11" onsubmit="return check();">
<input type="hidden" name="action" value="b_addsave">
	<tr>
		<td colspan="2" class="header"><b>商品大类添加</b></td>
	</tr>
	<tr>
		<td align="right">大类名称：</td>
		<td><input type="text" name="prod_BigClass_name" size="20"></td>
	</tr>
	<tr>
		<td align="right">大类排序：</td>
		<td><input type="text" name="prod_BigClass_sort" size="2"></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="保存设置" name="B5"></td>
	</tr>
</form>
</tbody>
</table>
<%
end sub

sub s_add()
%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Prod_Class_List.asp" method="post" name="form2" onsubmit="return check();">
<input type="hidden" name="action" value="s_addsave">
<input type="hidden" name="prod_SmallClass_bid" value="<%=request.querystring("prod_BigClass_id")%>">
	<tr>
		<td colspan="2" class="header"><b>商品小类添加</b></td>
	</tr>
	<tr>
		<td align="right">大类名称：</td>
		<td><img src="images/icon_arrow.gif">&nbsp; <b><%=request.querystring("prod_BigClass_name")%></b></td>
	</tr>
	<tr>
		<td align="right">现有小类：</td>
		<td>
		    <%
			 set rs=server.createobject("adodb.recordset")
			 sql="select prod_SmallClass_name from prod_SmallClass where prod_SmallClass_bid="&cint(request.querystring("prod_BigClass_id"))
			 rs.open sql,conn,1,1
			 if rs.eof then 
                 response.write "&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>此大类下暂无小类</font>"
             else
                 set prod_SmallClass_name=rs(0)
			     while not rs.eof
			     response.write "&nbsp;&nbsp;&nbsp;&nbsp;· "&prod_SmallClass_name&"<br>"
			     rs.movenext
			     wend
			 end if
			 rs.close
			 set rs=nothing
			%>
        </td>
	</tr>
	<tr>
		<td align="right">新增小类：</td>
		<td><input type="text" name="prod_SmallClass_name" size="20"></td>
	</tr>
	<tr>
		<td>
		<p align="right">小类排序：</td>
		<td><input type="text" name="prod_SmallClass_sort" size="2"></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="保存设置" name="B6"></td>
	</tr>
</form>
</tbody>
</table>
<%
end sub

sub b_modi()
prod_BigClass_id=my_request("prod_BigClass_id",1)
prod_BigClass_name=my_request("prod_BigClass_name",0)
prod_BigClass_sort=my_request("prod_BigClass_sort",1)
%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Prod_Class_List.asp" method="post" name="form4" onsubmit="return check();">
<input type="hidden" name="action" value="b_modisave">
<input type=hidden name=prod_BigClass_id value=<%=prod_BigClass_id%>>
	<tr>
		<td colspan="2" class="header"><b>商品大类修改</b></td>
	</tr>
	<tr>
		<td>
		<p align="right">大类名称：</td>
		<td><input type="text" name="prod_BigClass_name" size="20" value=<%=prod_BigClass_name%>></td>
	</tr>
	<tr>
		<td align="right">大类排序：</td>
		<td>
		<input type="text" name="prod_BigClass_sort" size="2" value="<%=prod_BigClass_sort%>"></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="保存设置" name="B7"></td>
	</tr>
</form>
</tbody>
</table>
<%
end sub
sub s_modi()
%>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Prod_Class_List.asp" method="post" name="form5" onsubmit="return check();">
<input type="hidden" name="action" value="s_modisave">
<input type="hidden" name="prod_SmallClass_id" value="<%=request.querystring("prod_SmallClass_id")%>">
	<tr>
		<td colspan="2" class="header"><b>商品小类修改</b></td>
	</tr>
	<tr>
		<td align="right">所属大类：</td>
		<td><b><%=request.querystring("prod_BigClass_name")%></b></td>
	</tr>
	<tr>
		<td align="right">小类名称：</td>
		<td><input type="text" name="prod_SmallClass_name" size="20" value=<%=request.querystring("prod_SmallClass_name")%>></td>
	</tr>
	<tr>
		<td>
		<p align="right">小类排序：</td>
		<td>
		<input type="text" name="prod_SmallClass_sort" size="2" value="<%=request.querystring("prod_SmallClass_sort")%>"></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="保存设置" name="B8"></td>
	</tr>
</form>
</tbody>
</table>
<%end sub%>
</body>

</html>
 
