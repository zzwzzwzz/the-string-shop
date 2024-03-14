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
dim username,password,userid,passid,user,pass
userid=checkStr(request.form("login_name"))
passid=checkStr(request.form("login_pass"))
user=replace(trim(userid),"'","")
pass=replace(trim(passid),"'","")
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
userip2 = Request.ServerVariables("REMOTE_ADDR")
if user="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�û�������Ϊ�գ�</li>"
end if
if pass="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>���벻��Ϊ�գ�</li>"
end if
if (instr(userid,"'")<>0 or instr(passid,"'")<>0) then 
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>��������¼��</li>"
end if

if FoundErr<>True then
	pass=md5(pass,32)
	set rs=server.createobject("adodb.recordset")
	sql="select * from admin_info where admin_info_PassWord='"&pass&"' and admin_info_UserName='"&user&"'"
	rs.open sql,conn,1,3
	    if rs.bof and rs.eof then
		    FoundErr=True
		    ErrMsg=ErrMsg & "<br><li>�û������������</li>"
	    else
	        if session("admin_info_UserName")=rs("admin_info_UserName") then
	            FoundErr=True
			    ErrMsg=ErrMsg & "<br><li><a href=adminloginout.asp>���û��Ѿ���¼��(����˳��������µ�¼)��</a></li>"
            else
		        if pass<>rs("admin_info_PassWord") then
			        FoundErr=True
			        ErrMsg=ErrMsg & "<br><li>�û������������</li>"
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
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub WriteErrMsg()
	dim strErr
    strErr=strErr & "<link rel=stylesheet type=text/css href=style.css>" & vbcrlf
    strErr=strErr & "<br><br><br><br><br><br><br><br><br><br><br><br>" & vbcrlf
    strErr=strErr & "<table cellspacing=1 cellpadding=5 width='20%' class=tableborder align=center>" & vbcrlf
    strErr=strErr & "<tbody class=altbg2>" & vbcrlf
    strErr=strErr & "	<tr><td class=header>������ʾ</td></tr>" & vbcrlf
    strErr=strErr & "	<tr><td><b>��������Ŀ���ԭ��</b><br>" & errmsg &"<br><br><a href='admin_login.asp'>&lt;&lt; ���ص�¼ҳ��</a></td></tr>" & vbcrlf
    strErr=strErr & "</table>" & vbcrlf
	response.write strErr
end sub
%>