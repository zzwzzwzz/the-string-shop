<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    bid			= my_request("bid",1)
    sid			= my_request("sid",1)
    product_info_name   = my_request("product_info_name",0)
    product_info_no   = my_request("product_info_no",0)
    product_info_flag   = my_request("product_info_flag",0)
    product_info_PriceM = my_request("product_info_PriceM",0)
    product_info_PriceS = my_request("product_info_PriceS",0)
    product_info_PicS   = my_request("product_info_PicS",0)
    product_info_PicB   = my_request("product_info_PicB",0)
    product_info_PicB2   = my_request("product_info_PicB2",0)
    product_info_PicB3   = my_request("product_info_PicB3",0)
    product_info_Detail = my_request("content",0)
    product_info_OnOff  = my_request("product_info_OnOff",0)
    product_info_KuCun  = my_request("product_info_KuCun",1)
    
    ErrMsg=""
    if product_info_name="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��Ʒ���Ʋ���Ϊ�գ�</li>"
    end if
    if bid="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��Ʒ��������ѡ��</li>"
    end if
    if sid="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��ƷС������ѡ��</li>"
    end if
    if product_info_PriceS="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��վ�۸���Ϊ�գ�</li>"
    end if

    if product_info_Detail="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��Ʒ��ϸ��������Ϊ�գ�</li>"
    end if
                     
    if FoundErr<>True then
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from product_info Where product_info_name='"&product_info_name&"'"
        rs.open sql,conn,1,1
      	if not rs.eof and rs.bof then
       		response.write "<script language='javascript'>"
        	response.write "alert('�����ˣ���Ʒ�����ظ���������¼�룡');"
        	response.write "location.href='javascript:history.go(-1)';"
        	response.write "</script>"
        	response.end
        end if
        rs.close
        set rs=nothing

        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from product_info"
        rs.open sql,conn,1,3
        rs.addnew
        rs("bid")=bid
        rs("sid")=sid
        rs("product_info_name")   = product_info_name
        rs("product_info_no")     = product_info_no
        rs("product_info_flag")   = product_info_flag
        rs("product_info_PriceM") = product_info_PriceM
        rs("product_info_PriceS") = product_info_PriceS
        rs("product_info_PicS")	  = product_info_PicS
        rs("product_info_PicB")   = product_info_PicB
        rs("product_info_PicB2")  = product_info_PicB2
        rs("product_info_PicB3")  = product_info_PicB3
        rs("product_info_Detail") = product_info_Detail
        rs("product_info_OnOff")  = product_info_OnOff
        rs("addtime")			  = now()
        rs("product_info_KuCun")  = product_info_KuCun
        rs.update
        rs.close
        set rs=nothing
        call ok("���ѳɹ������һ����Ʒ��Ϣ��","product_info_add.asp?bid="&bid&"&sid="&sid&"")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ʒ��Ϣ���</title>
<link rel="stylesheet" type="text/css" href="style.css">
<%
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from prod_SmallClass order by prod_SmallClass_id desc"
rs.open sql,conn,1,1
%>
<script language="JavaScript">
var onecount;
onecount=0;
subcat=new Array();
        <%
        count=0
        do while not rs.eof 
        %>
subcat[<%=count%>]=new Array("<%= trim(rs("prod_SmallClass_name"))%>","<%= trim(rs("prod_SmallClass_bid"))%>","<%= trim(rs("prod_SmallClass_id"))%>");
        <%
        count=count + 1
        rs.movenext
        loop
        rs.close
        set rs=nothing
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.form1.sid.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form1.sid.options[document.form1.sid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    } 
    
function showlist(dd)
   {
   if(dd.style.display=="none")
      {
        dd.style.display="";
      }
   else
      {
        dd.style.display="none";
      }
   }

</script>

</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Product_Info_Add.asp" method="post" name="form1">
<input type="hidden" name="action" value="save">
    <tr>
		<td colspan="2" class="title">��Ʒ��Ϣ���</td>
	</tr>
	<tr>
		<td>��Ʒ���ƣ�</td>
		<td><input type="text" name="product_info_name" size="30"></td>
	</tr>
	<tr>
		<td>��Ʒ���ţ�</td>
		<td><input type="text" name="product_info_no" size="30"></td>
	</tr>
	<tr>
		<td>�������</td>
		<td><select name="bid" onChange="changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)">
		    	<option>��ѡ�����</option>
		    	<%
		     	sql="select prod_BigClass_id,prod_BigClass_name from prod_BigClass order by prod_BigClass_id desc"
		     	set rs=conn.execute (sql)
		     	do while not rs.eof
		    	%>
		    	<option value="<%=rs("prod_BigClass_id")%>" <%if rs("prod_BigClass_id")=bid then response.write "selected" %>><%=rs("prod_BigClass_name")%></option>
		    	<%
		    	 rs.movenext
		     	loop
		     	rs.close
		     	set rs=nothing
		    	%>
		 	</select>
		 	<select name="sid">
		   		<option value="" <%if sid="" or null(sid) then response.write "selected" %>>��ѡ��С��</option>		  
           		<%if sid<>"" then%><option value="<%=prod_info_sid%>" selected><%=prod_SmallClass_name%></option><%end if%> 
         	</select>
		</td>
	</tr>

	<tr>
		<td>�� �� �ۣ�</td>
		<td><input type="text" name="product_info_PriceM" size="30"></td>
	</tr>
	<tr>
		<td>�� վ �ۣ�</td>
		<td><input type="text" name="product_info_PriceS" size="30"></td>
	</tr>
	<tr>
		<td>�� �� ����</td>
		<td>
		        <input type="text" name="product_info_KuCun" size="30">��</td>
	</tr>
	<tr>
		<td>С ͼ Ƭ��</td>
		<td>
		        <input type="text" name="product_info_PicS" size="30">
		        <input type="button" value=">>����ϴ���ƷСͼƬ" name="action0" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicS','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')">
		</td>
	</tr>
<tr>
		<td>�� ͼ Ƭ��</td>
		<td>
		        <input type="text" name="product_info_PicB" size="30">
		        <input type="button" value=">>����ϴ���Ʒ��ͼƬ" name="action1" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')">
        		<br>
		<input type="checkbox" name="MorePic" value="1" onClick='showlist(paipai);'>Ҫ�ϴ�������Ʒ��ͼƬ,����ǰ�淽���ڴ�(<font color="#808080">��๲֧��������Ʒ��ͼƬ</font>)</td>
	</tr>
	    <tr id=paipai style="display:none">
		<td>��ͼ�ϴ���</td>
		<td>
			�ڶ�����Ʒ��ͼ��<input type="text" name="product_info_PicB2" size="30" readonly> <a href="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB2','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=410,height=230')">
			<img src="images/upload.gif" alt="�ϴ�ͼƬ" style="cursor: hand;" onMouseOver="window.status='ʹ��ϵͳ�Դ����ϴ������ϴ�ͼƬ';return true;" onMouseOut="window.status='';return true;" border="0"></a><br>
			��������Ʒ��ͼ��<input type="text" name="product_info_PicB3" size="30" readonly> <a href="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB3','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=410,height=230')">
			<img src="images/upload.gif" alt="�ϴ�ͼƬ" style="cursor: hand;" onMouseOver="window.status='ʹ��ϵͳ�Դ����ϴ������ϴ�ͼƬ';return true;" onMouseOut="window.status='';return true;" border="0"></a>
		</td>
	</tr>
	<tr>
		<td>��ϸ������</td>
		<td>
		<textarea cols=60 rows=20 id="content" name="Content"></textarea>
		</td>
	</tr>
	<tr>
		<td>��Ʒ���ԣ�</td>
		<td><input type="checkbox" name="product_info_flag" value="1">��Ʒ&nbsp; 
		<input type="checkbox" name="product_info_flag" value="2">�Ƽ�&nbsp; 
		<input type="checkbox" name="product_info_flag" value="3">�ؼ�</td>
	</tr>
	<tr>
		<td>�Ƿ��ϼܣ�</td>
		<td><input type="radio" value="0" name="product_info_OnOff" checked>�ϼ�(��ʾ)&nbsp;&nbsp;
		<input type="radio" value="1" name="product_info_OnOff">�¼�(����) </td>
	</tr>
	<tr>
		<td>��</td>
		<td>
		<input type="submit" value="�ύ" name="Submit1">&nbsp;&nbsp;&nbsp; 
		<input type="reset" value="����" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>

