<%
response.buffer=true '启用缓冲处理
dim conn 
dim connstr
dim scadb
scadb=dbpath&"data/#data.mdb"   'dbpath 为各文件中设置的路径，请不要改动
on error resume next
connstr="DBQ="+server.mappath(""&scadb&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
If Err Then
    err.Clear
    Set Conn = Nothing
    Response.Write "数据库连接出错，请检查数据库连接文件中的数据库参数设置。"
    Response.End
End If 

Sub Chkhttp()
    Dim url1,url2
    url1=Cstr(Request.ServerVariables("HTTP_REFERER"))
    url2=Cstr(Request.ServerVariables("SERVER_NAME"))
    If mid(url1,8,len(url2))<>url2 Then
        Response.Write "参数错误"
        Response.End
    End If
    if instr(url1,"http://"&request.servervariables("host") )<1 then 
        response.write "处理 URL 时服务器上出错。如果您是在用任何手段攻击服务器，那你应该庆幸，你的所有操作已经被服务器记录，我们会第一时间通知公安局与国家安全部门来调查你的IP. "
        response.end
    end if
End Sub
%>
 

