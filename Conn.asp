<%
response.buffer=true '���û��崦��
dim conn 
dim connstr
dim scadb
scadb=dbpath&"data/#data.mdb"   'dbpath Ϊ���ļ������õ�·�����벻Ҫ�Ķ�
on error resume next
connstr="DBQ="+server.mappath(""&scadb&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
If Err Then
    err.Clear
    Set Conn = Nothing
    Response.Write "���ݿ����ӳ����������ݿ������ļ��е����ݿ�������á�"
    Response.End
End If 

Sub Chkhttp()
    Dim url1,url2
    url1=Cstr(Request.ServerVariables("HTTP_REFERER"))
    url2=Cstr(Request.ServerVariables("SERVER_NAME"))
    If mid(url1,8,len(url2))<>url2 Then
        Response.Write "��������"
        Response.End
    End If
    if instr(url1,"http://"&request.servervariables("host") )<1 then 
        response.write "���� URL ʱ�������ϳ���"
        response.end
    end if
End Sub
%>
 

