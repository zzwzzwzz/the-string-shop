<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=0
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_NumsPerRow,root_option_RowsPerPage,root_option_RowsIndexNew,root_option_RowsIndexTj,root_option_RowsIndexSpec,root_option_WidthSPic,root_option_HeighSPic,root_option_OnOffAliPayButton,root_option_GuestOrderOnOff,root_option_NumsPerRowSclass,root_option_NumsIndexHot from root_option where id=1"
rs.open sql,conn,1,1
root_option_NumsPerRow        = rs(0)
root_option_RowsPerPage       = rs(1)
root_option_RowsIndexNew      = rs(2)
root_option_RowsIndexTj       = rs(3)
root_option_RowsIndexSpec     = rs(4)
root_option_WidthSPic      	  = rs(5)
root_option_HeighSPic         = rs(6)
root_option_OnOffAliPayButton = rs(7)
root_option_GuestOrderOnOff=rs(8)
root_option_NumsPerRowSclass=rs(9)
root_option_NumsIndexHot=rs(10)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
   call save()
end if

sub save()
    root_option_NumsPerRowSclass  = my_request("root_option_NumsPerRowSclass",1)
    root_option_NumsPerRow        = my_request("root_option_NumsPerRow",1)
    root_option_RowsPerPage       = my_request("root_option_RowsPerPage",1)
    root_option_RowsIndexNew  	  = my_request("root_option_RowsIndexNew",1)
    root_option_RowsIndexTj 		  = my_request("root_option_RowsIndexTj",1)
    root_option_RowsIndexSpec      = my_request("root_option_RowsIndexSpec",1)
    root_option_WidthSPic         = my_request("root_option_WidthSPic",1)
    root_option_HeighSPic         = my_request("root_option_HeighSPic",1)
    root_option_OnOffAliPayButton = my_request("root_option_OnOffAliPayButton",1)
    root_option_GuestOrderOnOff   = my_request("root_option_GuestOrderOnOff",1)
	root_option_NumsIndexHot       =my_request("root_option_NumsIndexHot",1)
    ErrMsg=""
    if root_option_RowsPerPage="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>ÿҳ��ʾ��Ʒ��������Ϊ�գ�</li>"
    end if
    if root_option_WidthSPic="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��Ʒ��ͼ�ߴ�-�����Ϊ�գ�</li>"
    end if
    if root_option_HeighSPic="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��Ʒ��ͼ�ߴ�-���߲���Ϊ�գ�</li>"
    end if
    if FoundErr<>True then
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from root_option where id=1"
        rs.open sql,conn,1,3
        rs("root_option_NumsPerRowSclass")	= root_option_NumsPerRowSclass
        rs("root_option_NumsPerRow")     	= root_option_NumsPerRow
        rs("root_option_RowsPerPage")      	= root_option_RowsPerPage
        rs("root_option_RowsIndexNew")  	= root_option_RowsIndexNew
        rs("root_option_RowsIndexTj") 		= root_option_RowsIndexTj
        rs("root_option_RowsIndexSpec")	    = root_option_RowsIndexSpec
        rs("root_option_WidthSPic")         = root_option_WidthSPic
        rs("root_option_HeighSPic")         = root_option_HeighSPic
        rs("root_option_OnOffAliPayButton") = root_option_OnOffAliPayButton
        rs("root_option_GuestOrderOnOff")   = root_option_GuestOrderOnOff
        rs("root_option_NumsIndexHot")      = root_option_NumsIndexHot
        rs.update
        rs.close
        set rs=nothing

        call ok("���ѳɹ��������ѡ�����ã�","root_option_set.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ѡ��-����</title>
<link rel="stylesheet"  href="style.css" type="text/css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="root_option_set.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">����ѡ��-����</td>
	</tr>
	<tr>
		<td>��Ʒ������С���ÿ����ʾ��</td>
		<td><select size="1" name="root_option_NumsPerRowSclass">
		<option value="1" <%if cint(root_option_NumsPerRowSclass)=1 then response.write "selected"%>>1</option>
		<option value="2" <%if cint(root_option_NumsPerRowSclass)=2 then response.write "selected"%>>2</option>
		</select> ��</td>
	</tr>
	<tr>
		<td>ÿ����ʾ��Ʒ����</td>
		<td><select size="1" name="root_option_NumsPerRow">
		<option value="3" <%if cint(root_option_NumsPerRow)=3 then response.write "selected"%>>3</option>
		<option value="4" <%if cint(root_option_NumsPerRow)=4 then response.write "selected"%>>4</option>
		<option value="5" <%if cint(root_option_NumsPerRow)=5 then response.write "selected"%>>5</option>
		<option value="6" <%if cint(root_option_NumsPerRow)=6 then response.write "selected"%>>6</option>
		<option value="7" <%if cint(root_option_NumsPerRow)=7 then response.write "selected"%>>7</option>
		<option value="8" <%if cint(root_option_NumsPerRow)=8 then response.write "selected"%>>8</option>
		</select> ��</td>
	</tr>
	<tr>
		<td>ÿҳ��ʾ��Ʒ������</td>
		<td>
		<input type="text" name="root_option_RowsPerPage" size="3" value="<%=root_option_RowsPerPage%>"> ��(ֻ��������)</td>
	</tr>
	<tr>
		<td>��ҳ��Ʒ����-������</td>
		<td><select size="1" name="root_option_RowsIndexNew">
		<option value="1" <%if root_option_RowsIndexNew=1 then response.write "selected"%>>1</option>
		<option value="2" <%if root_option_RowsIndexNew=2 then response.write "selected"%>>2</option>
		<option value="3" <%if root_option_RowsIndexNew=3 then response.write "selected"%>>3</option>
		<option value="4" <%if root_option_RowsIndexNew=4 then response.write "selected"%>>4</option>
		</select> ��</td>
	</tr>
	<tr>
		<td>��ҳ�Ƽ���Ʒ-������</td>
		<td><select size="1" name="root_option_RowsIndexTj">
		<option value="1" <%if root_option_RowsIndexTj=1 then response.write "selected"%>>1</option>
		<option value="2" <%if root_option_RowsIndexTj=2 then response.write "selected"%>>2</option>
		<option value="3" <%if root_option_RowsIndexTj=3 then response.write "selected"%>>3</option>
		<option value="4" <%if root_option_RowsIndexTj=4 then response.write "selected"%>>4</option>
		</select> ��</td>
	</tr>
	<tr>
		<td>��ҳ�ؼ���Ʒ-������</td>
		<td><select size="1" name="root_option_RowsIndexSpec">
		<option value="1" <%if root_option_RowsIndexSpec=1 then response.write "selected"%>>1</option>
		<option value="2" <%if root_option_RowsIndexSpec=2 then response.write "selected"%>>2</option>
		<option value="3" <%if root_option_RowsIndexSpec=3 then response.write "selected"%>>3</option>
		<option value="4" <%if root_option_RowsIndexSpec=4 then response.write "selected"%>>4</option>
		</select> ��</td>
	</tr>
	<tr>
		<td>��ҳ-������Ʒ-��ʾ������</td>
		<td>
		<input type="text" name="root_option_NumsIndexHot" size="3" value="<%=root_option_NumsIndexHot%>">��</td>
	</tr>
	<tr>
		<td>��Ʒ��ͼ-��ʾ�ߴ�-���</td>
		<td><input type="text" name="root_option_WidthSPic" size="3" value="<%=root_option_WidthSPic%>"><font color="#BFBFBF">����&nbsp; </font>
		<font color="#BFBFBF">�����鲻Ҫ�����ģ�</font></td>
	</tr>
	<tr>
		<td>��Ʒ��ͼ-��ʾ�ߴ�-���ߣ�</td>
		<td>
		<input type="text" name="root_option_HeighSPic" size="3" value="<%=root_option_HeighSPic%>"><font color="#BFBFBF">����&nbsp; </font>
		<font color="#BFBFBF">�����鲻Ҫ�����ģ�</font></td>
	</tr>
	<tr>
		<td>֧�������أ�</td>
		<td>
		<input type="radio" value="1" name="root_option_OnOffAliPayButton" <%if root_option_OnOffAliPayButton=1 then response.write "checked"%>>����&nbsp;
		<input type="radio" value="0" name="root_option_OnOffAliPayButton" <%if root_option_OnOffAliPayButton=0 then response.write "checked"%>>�ر�</td>
	</tr>
	<tr>
		<td>�Ƿ�֧���ο��µ���</td>
		<td>
		<input type="radio" value="1" name="root_option_guestOrderOnOff" <%if root_option_GuestOrderOnOff=1 then response.write "checked"%>>ֻ֧��ע���Ա�µ�<br>
		<input type="radio" value="0" name="root_option_guestOrderOnOff" <%if root_option_GuestOrderOnOff=0 then response.write "checked"%>>ͬʱҲ֧���ο��µ�</td>
	</tr>
	<tr>
		<td>��</td>
		<td><input type="submit" value="  ��  ��  " name="B1">&nbsp;
		<input type="reset" value="  ��  ��  " name="B2"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
