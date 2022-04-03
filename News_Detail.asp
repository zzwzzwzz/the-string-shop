<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
id=my_request("id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""参数错误!"");location.href=""news_List.asp"";</script>")
  response.end
end if

//更新浏览次数
sql="update news_info set news_info_hitnums=news_info_hitnums+1 where id="&id
conn.execute (sql)

dim news_info_title,news_info_content,news_info_addtime,news_info_hitnums
sql="select news_info_title,news_info_content,news_info_addtime,news_info_hitnums from news_info where id="&id
set rs=conn.execute (sql)
news_info_title  =rs(0)
news_info_content=rs(1)
news_info_addtime=rs(2)
news_info_hitnums=rs(3)
rs.close
set rs=nothing

call up(news_info_title,"文章详情","<a href=News_List.asp>文章中心</a> &raquo; 文章详情")

response.write  "<tr>"&_
				"	<td><h2 align=center>"&news_info_title&"</h2></td>"&_
				"</tr>"&_
				"<tr><td align=center>发布时间："&news_info_addtime&"</font>&nbsp;&nbsp;浏览次数："&news_info_hitnums&"次</td></tr>"&_
				"<tr>"&_
				"	<td style='table-layout:fixed;word-break:break-all' class=maintxt>"&news_info_content&"<br></td>"&_
				"</tr>"
call down()
%></center>