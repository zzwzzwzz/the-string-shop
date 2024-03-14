<%
Function my_request(ParaName,ParaType)
    Dim ParaValue
    ParaValue=Request(ParaName)
    If ParaType=1 Then
        ErrMsg=""
        If Not isNumeric(ParaValue) Then
            FoundErr=True
	        ErrMsg="<li>非法操作！</li>"
	        call WriteErrMsg(ErrMsg)
            response.end
        end if
    Else
        ParaValue=replace(ParaValue,"'","''")
    End if
    my_request=ParaValue
End function

'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub WriteErrMsg1(ErrMsg)
	dim strErr
	strErr=strErr & "<br><br><table width='40%' align=center cellpadding=6 style='border: 1px solid #CCCCCC; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px' cellspacing=6>"
	strErr=strErr & "<tr>"
	strErr=strErr & "<td style='border-bottom: 1px dashed #CCCCCC;' bgcolor=#EBEBEB>"
	strErr=strErr & "<font color=#FF0000><b><span style='font-size: 14px'>错误提示：</span></b></font></td>"
	strErr=strErr & "</tr>"
	strErr=strErr & "<tr>"
	strErr=strErr & "<td><span style='font-size: 12px'><b>产生错误的可能原因：</b><br>" & ErrMsg &"<br><br><a href='javascript:history.go(-1)'>&lt;&lt; 点此返回操作</a></span></td>"
	strErr=strErr & "</tr>"
	strErr=strErr & "</table>"
	response.write strErr
	response.end
end sub


'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub WriteErrMsg(ErrMsg)
	dim strErr
    call up("错误提示","错误提示","错误提示")
	strErr=strErr & "<tr><td>"&ErrMsg&"</td></tr>"
	strErr=strErr & "<tr><td align=center><b><a href='javascript:history.go(-1)'>&lt;&lt; 点此返回操作</a></b></td></tr>"
	response.write strErr
    call down()
    response.end
end sub

'****************************************************
'过程名：WriteErrMsg2
'作  用：显示错误提示信息(后台用)
'参  数：无
'****************************************************
sub WriteErrMsg2(ErrMsg)
	dim strErr
	strErr=strErr & "<link rel=stylesheet type=text/css href=style.css>"
	strErr=strErr & "<br><br><table cellspacing=1 cellpadding=4 width=40% class=tableborder align=center>"
	strErr=strErr & "<tbody class=altbg2>"
	strErr=strErr & "    <tr><td class=title>错误提示：</td></tr>"
	strErr=strErr & "	<tr><td>"&ErrMsg&"</td></tr>"
	strErr=strErr & "	<tr><td align=center><a href='javascript:history.go(-1)'>&lt;&lt; 点此返回操作</a></td></tr>"
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
  Response.write "		<td class=header>操作成功提示</td>"
  Response.write "	</tr>"
  Response.write "	<tr>"
  Response.write "		<td align=center><font color=red>:) 恭喜，您进行的操作已经成功完成。正在转向中,请稍候......</font></td>"
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
    response.write "<SCRIPT language=JavaScript>alert('出错了，资料填写不完整或者不符合要求，请检查后重新提交。');"
    response.write "location.href='javascript:history.go(-1)';</SCRIPT>"
    response.end
end sub
%>


