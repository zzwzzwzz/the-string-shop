<%
Function my_request(ParaName,ParaType)
    Dim ParaValue
    ParaValue=Request(ParaName)
    If ParaType=1 Then
        ErrMsg=""
        If Not isNumeric(ParaValue) Then
            FoundErr=True
	        ErrMsg="<li>�Ƿ�������</li>"
	        call WriteErrMsg(ErrMsg)
            response.end
        end if
    Else
        ParaValue=replace(ParaValue,"'","''")
    End if
    my_request=ParaValue
End function

'****************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub WriteErrMsg1(ErrMsg)
	dim strErr
	strErr=strErr & "<br><br><table width='40%' align=center cellpadding=6 style='border: 1px solid #CCCCCC; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px' cellspacing=6>"
	strErr=strErr & "<tr>"
	strErr=strErr & "<td style='border-bottom: 1px dashed #CCCCCC;' bgcolor=#EBEBEB>"
	strErr=strErr & "<font color=#FF0000><b><span style='font-size: 14px'>������ʾ��</span></b></font></td>"
	strErr=strErr & "</tr>"
	strErr=strErr & "<tr>"
	strErr=strErr & "<td><span style='font-size: 12px'><b>��������Ŀ���ԭ��</b><br>" & ErrMsg &"<br><br><a href='javascript:history.go(-1)'>&lt;&lt; ��˷��ز���</a></span></td>"
	strErr=strErr & "</tr>"
	strErr=strErr & "</table>"
	response.write strErr
	response.end
end sub


'****************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub WriteErrMsg(ErrMsg)
	dim strErr
    call up("������ʾ","������ʾ","������ʾ")
	strErr=strErr & "<tr><td>"&ErrMsg&"</td></tr>"
	strErr=strErr & "<tr><td align=center><b><a href='javascript:history.go(-1)'>&lt;&lt; ��˷��ز���</a></b></td></tr>"
	response.write strErr
    call down()
    response.end
end sub

'****************************************************
'��������WriteErrMsg2
'��  �ã���ʾ������ʾ��Ϣ(��̨��)
'��  ������
'****************************************************
sub WriteErrMsg2(ErrMsg)
	dim strErr
	strErr=strErr & "<link rel=stylesheet type=text/css href=style.css>"
	strErr=strErr & "<br><br><table cellspacing=1 cellpadding=4 width=40% class=tableborder align=center>"
	strErr=strErr & "<tbody class=altbg2>"
	strErr=strErr & "    <tr><td class=title>������ʾ��</td></tr>"
	strErr=strErr & "	<tr><td>"&ErrMsg&"</td></tr>"
	strErr=strErr & "	<tr><td align=center><a href='javascript:history.go(-1)'>&lt;&lt; ��˷��ز���</a></td></tr>"
	strErr=strErr & "</tbody>"
	strErr=strErr & "</table>"
	response.write strErr
	response.end
end sub

Sub SaveOk(url)
  Response.write "<link rel=stylesheet type=text/css href=../Admin/style.css>"
  Response.write "<br><br><br><br><br>"
  Response.write "<table cellspacing=1 cellpadding=4 width='50%' class=tableborder align=center>"
  Response.write "<tbody class=altbg2>"
  Response.write "	<tr>"
  Response.write "		<td class=header>�����ɹ���ʾ</td>"
  Response.write "	</tr>"
  Response.write "	<tr>"
  Response.write "		<td align=center><font color=red>:) ��ϲ�������еĲ����Ѿ��ɹ���ɡ�����ת����,���Ժ�......</font></td>"
  Response.write "	</tr>"
  Response.write "</table>"
  Response.write "<meta http-equiv=""refresh"" content=""2;url="&url&""">"
  Response.end
End Sub

sub ok(txt,url)
    response.write "<script language='javascript'>"
    response.write "alert('"&txt&"');"
    response.write "location.href='"&url&"';"
    response.write "</script>"
end sub

sub error()
    response.write "<SCRIPT language=JavaScript>alert('�����ˣ�������д���������߲�����Ҫ������������ύ��');"
    response.write "location.href='javascript:history.go(-1)';</SCRIPT>"
    response.end
end sub
%>


