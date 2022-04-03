<%dim dbpath
dbpath="../"
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../include/char.asp"-->
<!--#include file="../include/md5.asp"-->
<!--#include file="../include/DuoDuoCode.asp"-->
<%
Call Chkhttp()
dim sql,rs
dim username,password,CheckCode,userid,passid,codeid,user,pass
userid=checkStr(request.form("login_name"))
passid=checkStr(request.form("login_pass"))
codeid=checkStr(request.form("codeid"))
user=replace(trim(userid),"'","")
pass=replace(trim(passid),"'","")
CheckCode=replace(trim(request.form("codeid")),"'","")
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
userip2 = Request.ServerVariables("REMOTE_ADDR")
if user="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>用户名不能为空！</li>"
end if
if pass="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>密码不能为空！</li>"
end if
if CheckCode="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>验证码不能为空！</li>"
end if
if session("CheckCode")="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>你登录时间过长，请重新返回登录页面进行登录。</li>"
end if
if CheckCode<>CStr(session("CheckCode")) then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>您输入的确认码和系统产生的不一致，请重新输入。</li>"
end if
if (instr(userid,"'")<>0 or instr(passid,"'")<>0) then 
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>非正常登录！</li>"
end if

if FoundErr<>True then
	pass=md5(pass,32)
	set rs=server.createobject("adodb.recordset")
	sql="select * from admin_info where admin_info_PassWord='"&pass&"' and admin_info_UserName='"&user&"'"
	rs.open sql,conn,1,3
	    if rs.bof and rs.eof then
		    FoundErr=True
		    ErrMsg=ErrMsg & "<br><li>用户名或密码错误！</li>"
	    else
	        if session("admin_info_UserName")=rs("admin_info_UserName") then
	            FoundErr=True
			    ErrMsg=ErrMsg & "<br><li><a href=adminloginout.asp>此用户已经登录过(点击退出后再重新登录)！</a></li>"
            else
		        if pass<>rs("admin_info_PassWord") then
			        FoundErr=True
			        ErrMsg=ErrMsg & "<br><li>用户名或密码错误！</li>"
		        else
                    session("admin_info_RealName")=rs("admin_info_RealName")
                    session("admin_info_UserName")=rs("admin_info_UserName")
                    Session("pass")=true
                    rs.close
                    set rs=nothing
                    conn.close
                    set conn=nothing
                    response.redirect "Index.asp"
                end if
            end if
        end if
    rs.close
    set rs=nothing
end if

if FoundErr=True then
   call WriteErrMsg()
end if

conn.close
set conn=nothing
   
'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub WriteErrMsg()
	dim strErr
    strErr=strErr & "<link rel=stylesheet type=text/css href=style.css>" & vbcrlf
    strErr=strErr & "<br><br><br><br><br>" & vbcrlf
    strErr=strErr & "<table cellspacing=1 cellpadding=4 width='30%' class=tableborder align=center>" & vbcrlf
    strErr=strErr & "<tbody class=altbg2>" & vbcrlf
    strErr=strErr & "	<tr><td class=header>错误提示</td></tr>" & vbcrlf
    strErr=strErr & "	<tr><td><b>产生错误的可能原因：</b><br>" & errmsg &"<br><br><a href='admin_login.asp'>&lt;&lt; 返回登录页面</a></td></tr>" & vbcrlf
    strErr=strErr & "</table>" & vbcrlf
	response.write strErr
end sub
%>

 
