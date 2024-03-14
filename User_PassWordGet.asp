<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<!--#include file="include/md5.asp"-->
<%
call up("取回密码","取回密码","取回密码")
action=my_request("action",0)
select case action
    case ""
        call getuser()
    case "setp1"
        call getquestion()
    case "setp2"
        call dpassok()
    case "setp3"
        call isok()
    case else
        call getuser()
end select


response.write "<tr><td colspan=2>"

sub getuser() '第一步：用户名验证%>
<table cellpadding="4" width="100%" class="tableborder" style="border-collapse: collapse">
<tbody class="table_td">
<form action="user_PassWordGet.asp" method="post" name=form1 onsubmit="return chkinput();">
<input name=action type=hidden value=setp1>
	<tr>
		<td width="100%" colspan="2" class="altbg1">
		<div id="Content">
			<div class="ForgotPassword" id="Forgot">
				<p class="WarningMsg">第1步、请输入您的会员名</div>
		</div>
		</td>
	</tr>
	<tr>
		<td width="36%">
		<p align="right"><span style="font-size: 12px">会员名：</span></td>
		<td width="62%"><span style="font-size: 12px"><input type="text" name="email" size="20" maxlength="20">
		</span>
		</td>
	</tr>
	<tr>
		<td width="36%">　</td>
		<td width="62%"><span style="font-size: 12px"><input class=button type="submit" value="下一步" name="B1"></span></td>
	</tr>
</form>
</tbody>
</table>

<%end sub

sub getquestion() '第二步：问题与答案验证
    email=my_request("email",0)
    if email="" then
        response.write"<SCRIPT language=JavaScript>alert('请输入用户名！');"
        response.write"javascript:history.go(-1)</SCRIPT>"
        response.end
    end if
    set rs=server.createobject("adodb.recordset")
    sql="select User_info_question from User_info where User_info_userName='"&email&"'"
    rs.open sql,conn,1,3
    if rs.eof then
        txt="<tr><td colspan=2 align=center><br><li><a href=""#"" onclick=history.back()>此用户不存在,请返回！</a></li><br><br></td></tr>"
    else
        question=rs(0)
    end if
    rs.close
    set rs=nothing
%>
<table cellpadding="4" width="100%" class="tableborder" style="border-collapse: collapse">
<tbody class="table_td">
  <form action="user_PassWordGet.asp" method="post" name=form1 onsubmit="return chkinput();">
  <input name=action type=hidden id=action value=setp2>
  <input name=email type=hidden id=email value="<%=email%>">
  <input name=question type=hidden id=question value="<%=question%>">        
	<tr>
		<td width="96%" colspan="2" class="altbg1">
		<div id="Content0">
			<div class="ForgotPassword" id="Forgot0">
				<p class="WarningMsg">第2步、请输入您的密码取回问题的答案 </div>
		</div>
		</td>
    </tr>
    <%
	if txt<>"" then 
	    response.write txt 
	else
	%>
	<tr>
		<td width="36%" align="right"><span style="font-size: 12px">您的问题：</td>
        <td width="60%" height="20"><span style="font-size: 12px"><%=question%></span></td>
    </tr>
	<tr>
		<td width="36%" align="right"><span style="font-size: 12px">请输入答案：</span></td>
		<td width="62%">
         <span style="font-size: 12px"><input type="text" name="answer" size="20" maxlength="20"></span>
        </td>
	</tr>
	<tr>
		<td width="36%">　</td>
		<td width="62%"><span style="font-size: 12px">
		<input class=button type="submit" value="下一步" name="B2"></span></td>
	</tr>
	<%end if%>
</form>
</tbody>
</table>
<%end sub

sub dpassok() '重新设置新密码
    email=my_request("email",0)
    question=my_request("question",0)
    answer=my_request("answer",0)
    set rs=server.createobject("adodb.recordset")
    sql="select * from User_info where User_info_userName='"&email&"' and User_info_answer='"&answer&"' and User_info_question='"&question&"'"
    rs.open sql,conn,1,3
    if rs.eof then
        txt="<tr><td colspan=2 align=center><br><li><a href=""#"" onclick=history.back()>您输入的答案有误，请返回！</a></li><br><br></td></tr>"
    end if
    rs.close
    set rs=nothing
%>
<table cellpadding="4" width="100%" class="tableborder" style="border-collapse: collapse">
<tbody class="table_td">
  <form action="user_PassWordGet.asp" method="post" name=form1 onsubmit="return chkinput();">
  <input name=action type=hidden id=action value=setp3>
  <input name=email type=hidden id=email value="<%=email%>">
    <tr>
       <td width="100%" align="right" height="20" colspan="2" class="altbg1">
		<p align="left">第3步、请重新设置新密码</td>
	</tr>
	<%
	if txt<>"" then 
	    response.write txt 
	else
	%>
    <tr>
       <td width="36%" align="right" height="20">
		<p><span style="font-size: 12px">请输入新密码：</span></span></td>
		<td width="62%"><span style="font-size: 12px">
		<input type=password name=password size=12 maxlength=15>(5-15位)</span></td>
        </span></td>
	</tr>
	<tr>
		<td width="36%" align="right"><span style="font-size: 12px">请确认新密码：</span></td>
		<td width="62%">
           <span style="font-size: 12px">
           <input type=password name=password2 size=12 maxlength=15>(5-15位)
        	</span>
        </td>
    </tr>
	<tr>
		<td width="36%">　</td>
		<td width="62%"><span style="font-size: 12px"><input class=button type="submit" value="下一步" name="B3"></span></td>
	</tr>
	<%end if%>
</form>
</tbody>
</table>
<%end sub

sub ok() '置新密码成功返回页
%>
<table cellpadding="4" width="100%" class="tableborder" style="border-collapse: collapse">
<tbody class="table_td">
	<tr>
		<td align="center" >新密码已设置成功!</td>
    </tr>
</tbody>
</table>
<%end sub

sub isok()
    password=my_request("password",0)
    password2=my_request("password2",0)
    email=my_request("email",0)
    if password="" or password2="" then
        response.write"<SCRIPT language=JavaScript>alert('请填写密码！');"
        response.write"javascript:history.go(-1)</SCRIPT>"
        response.end
    end if
    if password<>password2 then
        response.write"<SCRIPT language=JavaScript>alert('两次密码输入不一致！');"
        response.write"javascript:history.go(-1)</SCRIPT>"
        response.end
    end if
    password=md5(password,32)
    sql="update User_info set User_info_passWord='"&password&"' where User_info_userName='"&email&"'"
    conn.execute (sql)
    Response.write "<script>alert(""密码修改成功，请用新的密码登录！"");</script>"
    call ok()
end sub

response.write "</td></tr>"
call down()
%></center>