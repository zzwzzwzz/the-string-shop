<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_info_OnOff,root_info_OffNote,root_info_LogoPic,root_info_ICP,root_info_tel,root_info_email,root_info_QQ,root_info_QQOnOff,root_info_MSN,root_info_skin,root_info_IndexTitle,root_info_IndexKeyWords,root_info_IndexDescription,root_info_WangWangOnOff,root_info_WangWang,root_info_sitename,root_info_address,root_info_zip,root_info_fax from root_info where id=1"
rs.open sql,conn,1,1
root_info_OnOff           =rs(0)
root_info_OffNote         =rs(1)
root_info_LogoPic         =rs(2)
root_info_ICP             =rs(3)
root_info_tel             =rs(4)
root_info_email           =rs(5)
root_info_QQ              =rs(6)
root_info_QQOnOff         =rs(7)
root_info_MSN             =rs(8)
root_info_skin            =rs(9)
root_info_IndexTitle      =rs(10)
root_info_IndexKeyWords   =rs(11)
root_info_IndexDescription=rs(12)
root_info_WangWangOnOff   =rs(13)
root_info_WangWang        =rs(14)
root_info_sitename 		  =rs(15)
root_info_address 		  =rs(16)
root_info_zip			  =rs(17)
root_info_fax			  =rs(18)
root_info_QQPlace         =rs(19)
rs.close
set rs=nothing

if root_info_skin="" then
    response.write "<link href=style/001.css rel=stylesheet type=text/css>"
else
    response.write "<link href=style/"&root_info_skin&".css rel=stylesheet type=text/css>"
end if

if root_info_OnOff=1 then 
    response.write "<center><br><br><br><br><br>"&root_info_OffNote&"</center>"
    response.end
end if
%>


