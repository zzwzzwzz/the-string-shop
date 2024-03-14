<center><%
dim dbpath
dbpath=""
%>
<!--#include file="conn.asp"-->
<!--#include file="include/char.asp"-->
<!--#include file="include/MyRequest.asp"-->
<!--#include file="include/md5.asp"-->
<!--#include file="sub.asp"-->
<%
Call Chkhttp()
dim sql,rs
dim username,password,userid,passid

userid=checkStr(request.form("loginname"))
passid=checkStr(request.form("loginpass"))
user=replace(trim(userid),"'","")
pass=replace(trim(passid),"'","")
urlpath=my_request("urlpath",0)

ErrMsg=""
if user="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<li>用户名不能为空！</li>"
end if
if pass="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<li>密码不能为空！</li>"
end if
if FoundErr<>True then
	pass=md5(pass,32)
	set rs=server.createobject("adodb.recordset")
	sql="select * from user_info where user_info_PassWord='"&pass&"' and user_info_UserName='"&user&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
	    response.write"<SCRIPT language=JavaScript>alert('用户名或密码错误！');"
        response.write"javascript:history.go(-1)</SCRIPT>"
        response.end
	else
	    if session("user_info_UserName")=rs("user_info_UserName") then
	        response.write"<SCRIPT language=JavaScript>alert('此用户已经登录！');"
            response.write"javascript:history.go(-1)</SCRIPT>"
            response.end	        
        else
		    if pass<>rs("user_info_PassWord") then
	            response.write"<SCRIPT language=JavaScript>alert('用户名或密码错误！');"
                response.write"javascript:history.go(-1)</SCRIPT>"
                response.end
		    else
		    	if rs("user_info_states")=1 then 
		    		response.write"<SCRIPT language=JavaScript>alert('会员状态被锁定或审核中！');"
                	response.write"javascript:history.go(-1)</SCRIPT>"
                	response.end
                else
		        	rs("user_info_LastLoginTime")=now()
                	rs("user_info_LoginNums")=rs("user_info_LoginNums")+1
                	rs.update
                	session("user_info_id")=rs("user_info_id")
                	session("user_info_UserName")=rs("user_info_UserName")
                	session("user_info_LoginIn")=true
                	rs.close
                	set rs=nothing
                	Session.Timeout=30
                	if urlpath<>"" then
                    	response.redirect urlpath
                	else
                    	response.redirect "User_Index.asp"
                	end if
                end if
            end if
        end if
    end if
	rs.close
	set rs=nothing
else
	call WriteErrMsg(ErrMsg)
end if
conn.close
set conn=nothing
%>
 
</center>