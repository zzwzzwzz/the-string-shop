<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<!--#include file=include/Pages.asp -->
<%
call up("��������","��������","��������")

set rs=server.createobject("adodb.recordset")
sql="select id,news_info_title,news_info_addtime,news_info_content from news_info order by id desc"
rs.open sql,conn,1,1
if rs.eof then 
    response.write "<tr><td align=center>����������Ϣ!</td></tr>"
else
    rs.PageSize =20 'ÿҳ��¼����
    iCount=rs.RecordCount '��¼����
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
  	
  	dim id,news_info_title,news_info_addtime
  	set id                = rs(0)
  	set news_info_title   = rs(1)
  	set news_info_addtime = rs(2)                    
    while not rs.eof and i<=rs.pagesize
                    			
    txt="<a href=News_Detail.asp?id="&id&">"&news_info_title&"</a>"

  	response.write "<tr><td class=maintxt>"&txt&" ("&datevalue(news_info_addtime)&")</td></tr>"
    rs.movenext
    i=i+1
    wend
    call PageControl(iCount,maxpage,page)
end if
rs.close
set rs=nothing

call down()
%>
</center>