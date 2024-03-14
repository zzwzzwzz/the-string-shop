<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file="include/Pages.asp"-->
<!--#include file=Sub.asp -->
<%
action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    guest_info_name  =my_request("guest_info_name",0)
    guest_info_email =my_request("guest_info_email",0)
    guest_info_detail=my_request("guest_info_detail",0)

    ErrMsg=""
    if guest_info_name="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>名字不能为空！</li>"
    end if
    if guest_info_email="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>Email不能为空！</li>"
    end if
    if guest_info_detail="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>留言内容不能为空！</li>"
    end if

    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from guest_info"
        rs.open sql,conn,1,3
        rs.addnew
        rs("guest_info_name")  =guest_info_name
        rs("guest_info_email") =guest_info_email
        rs("guest_info_detail")=guest_info_detail
        rs("guest_info_time")  =now()
        rs.update
        rs.close
        set rs=nothing
        call ok("恭喜，您已成功添加新留言！","guestbook_List.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub

call up("在线留言","在线留言","在线留言")
    			set rs=server.createobject("adodb.recordset")
    			sql="select guest_info_name,guest_info_email,guest_info_detail,guest_info_time,guest_info_backdetail,guest_info_backTime from guest_info order by guest_info_id desc"
    			rs.open sql,conn,1,1
    			if rs.eof then 
    				response.write "<tr><td align=center colspan=2>目前暂无留言信息!</td></tr>"
    			else
    				rs.PageSize =20 '每页记录条数
        			iCount=rs.RecordCount '记录总数
        			iPageSize=rs.PageSize
        			maxpage=rs.PageCount 
        			page=request("page")  
        			if Not IsNumeric(page) or page="" then
        			    page=1
        			else
         			    page=cint(page)
        		    end if    
        			if page<1 then
        			    page=1
        			elseif  page>maxpage then
        			    page=maxpage
        			end if   
        			rs.AbsolutePage=Page
        			if page=maxpage then
	    			    x=iCount-(maxpage-1)*iPageSize
        			else
	    			    x=iPageSize
        			end if
        			i=1
        
        			set guest_info_name      =rs(0)
        		    set guest_info_email     =rs(1)
        			set guest_info_detail    =rs(2)
        			set guest_info_time      =rs(3)
        			set guest_info_backdetail=rs(4)
        			set guest_info_backTime  =rs(5)
        
        			while not rs.eof and i<=rs.pagesize
response.write  "<tr>"&_
				"	<td valign=top>"&guest_info_name&"</td>"&_
				"	<td valign=top>"&guest_info_detail&"&nbsp;&nbsp;("&guest_info_time&")<br>"
						if guest_info_backdetail<>"" then
response.write "			<b><font color=#EECCCC>管理员回复:</b></font><font color=#EECCCC>"&guest_info_backdetail&"</font>"
						end if
response.write  "	</td>"&_
				"</tr>"
    				rs.movenext
    				i=i+1
    				wend
    				response.write "<tr><td colspan=2>"
    				call PageControl(iCount,maxpage,page)
    				response.write "</td></tr>"
    			end if
    			rs.close
    			set rs=nothing
    			
response.write  "<form action=guestbook_List.asp method=post name=form1>"&_
				"<input type=hidden name=action value=save>"&_
				"<tr><td colspan=2 class=RightHead><a name=add>发表留言：</a></td></tr>"
				if session("user_info_LoginIn")=true then
					Set rs= Server.CreateObject("ADODB.Recordset")
					sql="select user_info_RealName,user_info_email from user_info where user_info_id="&session("user_info_id")
					rs.open sql,conn,1,1
					user_info_RealName=rs(0)
					user_info_email=rs(1)
					rs.close
					set rs=nothing
response.write  "	<tr><td>姓名：</td><td><input type=text name=guest_info_name size=30 value="&user_info_RealName&"></td></tr>"&_
				"	<tr><td>Email：</td><td><input type=text name=guest_info_email size=30 value="&user_info_email&"></td></tr>"
				else
response.write  "	<tr><td>姓名：</td><td><input type=text name=guest_info_name size=30></td></tr>"&_
				"	<tr><td>Email：</td><td><input type=text name=guest_info_email size=30></td></tr>"
				end if
response.write  "<tr><td valign=top>留言：</td><td><textarea rows=10 name=guest_info_detail cols=50></textarea></td></tr>"&_
				"<tr><td> </td><td><input type=submit value= 提交留言 ></td></tr>"&_
				"</form>"
call down()
%>
</center>