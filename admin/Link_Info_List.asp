<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=8
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/Pages.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>友情链接-管理</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
<script language = "JavaScript">   
//全选操作    
function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }

</script>
<%
action=my_request("action",0)
if action="删除" then
    call del()
end if

//过程：批量删除
sub del()
    id=my_request("id",0)
    if id<>"" then
       pp=ubound(split(id,","))+1 '判断数组id中共有几维
       for v=1 to pp
          id=request("id")(v)

          
          sql="select link_info_detail,link_info_type from link_info where id="&id
          set rs=conn.execute (sql)
          link_info_detail=rs("link_info_detail")
          link_info_type=rs("link_info_type")
          rs.close
          set rs=nothing
                   
          conn.execute ("delete from [link_info] where id="&id)
          
          //删除相应logo图片
          if link_info_type=1 then
              Dbpath="../uploadpic/"&link_info_detail
              Dbpath=server.mappath(Dbpath)
              bkfolder="../uploadpic"
              Set Fso=server.createobject("scripting.filesystemobject")
              if fso.fileexists(dbpath) then
                  If CheckDir(bkfolder) = True Then
                      fso.DeleteFile dbpath
                  end if
              end if
              Set fso = nothing
          end if

       next

      call ok("所选信息已成功删除！","link_info_list.asp")
    end if
end sub

Function CheckDir(FolderPath)
    folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
        CheckDir = True
    Else
        CheckDir = False
    End if
    Set fso1 = nothing
End Function
%>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td colspan="6" class="header">友情链接-管理</td>
	</tr>
	<tr>
		<td class="altbg1">选中</td>
		<td class="altbg1">类型</td>
		<td class="altbg1">链接文本/图标</td>
		<td class="altbg1">链接网址</td>
		<td class="altbg1">是否首页显示</td>
		<td class="altbg1" align="center">修改</td>
	</tr>
    <form name="form1" action="Link_info_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    sql="select * from link_info order by id desc"
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=6 align=center>目前暂无友情链接信息,<a href=link_info_add.asp>请添加!</a></td></tr>"
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
        while not rs.eof and i<=rs.pagesize
    %>
	<tr>
		<td><input type="checkbox" name="id" value="<%=rs("id")%>"></td>
		<td><%if rs("link_info_type")=1 then response.write "图标链接" else response.write "文字链接"%></td>
		<td><%if rs("link_info_type")=1 then response.write "<img src=../uploadpic/"&rs("link_info_detail")&">" else response.write rs("link_info_detail")%></td>
		<td><a href="<%=rs("link_info_url")%>" target=_blank><%=rs("link_info_url")%></a></td>
		<td><%if rs("link_info_IndexShow")=1 then response.write "否" else response.write "是"%></td>
		<td align="center"><a href="link_info_modi.asp?id=<%=rs("id")%>">修改</a></td>
	</tr>
	<%
         rs.movenext
         i=i+1
     wend
    %>
	<tr>
		<td colspan="6">
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选 
        <input type="submit" name="action" value="删除" onclick="{if(confirm('删除后将无法恢复，您确定要删除选定的信息吗？')){this.document.form1.submit();return true;}return false;}">&nbsp;
		<input type="button" value="添加友情链接" name="action1" onclick="window.location='link_info_add.asp'"></td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    <%
        call PageControl(iCount,maxpage,page)
    end if
    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
    %>
    </form>
</tbody>
</table>


</body>

</html>

