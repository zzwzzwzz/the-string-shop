<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""��������!"");location.href=""product_info_List.asp"";</script>")
  response.end
end if

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select id,bid,sid,product_info_name,product_info_flag,product_info_PriceM,product_info_PriceS,product_info_PicB,product_info_PicB2,product_info_PicB3,product_info_PicS,product_info_OnOff,product_info_KuCun,product_info_no,product_info_Detail from product_info where id="&id
rs.open sql,conn,1,1
id					=rs(0)
bid					=rs(1)
sid					=rs(2)
product_info_name	=rs(3)
product_info_flag	=rs(4)
product_info_PriceM	=rs(5)
product_info_PriceS	=rs(6)
product_info_PicB	=rs(7)
product_info_PicB2	=rs(8)
product_info_PicB3	=rs(9)
product_info_PicS	=rs(10)
product_info_OnOff  =rs(11)
product_info_KuCun  =rs(12)
product_info_no  	=rs(13)
product_info_Detail  =rs(14)
rs.close
set rs=nothing

sql="select prod_SmallClass_name from prod_SmallClass where prod_SmallClass_id="&Sid
set rs=conn.execute (sql)
SClass1=rs("prod_SmallClass_name")
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    id					= my_request("id",1)
    bid					= my_request("bid",1)
    sid					= my_request("sid",1)
    product_info_name   = my_request("product_info_name",0)
    product_info_flag   = my_request("product_info_flag",0)
    product_info_PriceM = my_request("product_info_PriceM",0)
    product_info_PriceS = my_request("product_info_PriceS",0)
    product_info_PicS   = my_request("product_info_PicS",0)
    product_info_PicB   = my_request("product_info_PicB",0)
    product_info_PicB2  = my_request("product_info_PicB2",0)
    product_info_PicB3  = my_request("product_info_PicB3",0)
    product_info_Detail = my_request("Content",0)
    product_info_OnOff  = my_request("product_info_OnOff",1)
    product_info_KuCun  = my_request("product_info_KuCun",1)
    product_info_no  	= my_request("product_info_no",0)
    
    ErrMsg=""
    if id="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��ƷID����Ϊ�գ�</li>"
    end if
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
        set rs=server.createobject("adodb.recordset")
        sql="select * from product_info Where product_info_name='"&product_info_name&"' and id<>"&id
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

		set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from product_info where id="&id
        rs.open sql,conn,1,3
        rs("bid")=bid
        rs("sid")=sid
        rs("product_info_name")   = product_info_name
        rs("product_info_no")     = product_info_no
        rs("product_info_flag")   = product_info_flag
        rs("product_info_PriceM") = product_info_PriceM
        rs("product_info_PriceS") = product_info_PriceS
        rs("product_info_PicB")   = product_info_PicB
        rs("product_info_PicB2")  = product_info_PicB2
        rs("product_info_PicB3")  = product_info_PicB3
        rs("product_info_PicS")	  = product_info_PicS
        rs("product_info_Detail") = product_info_Detail
        rs("product_info_OnOff")  = product_info_OnOff
        rs("addtime")			  = now()
        rs("product_info_KuCun")  = product_info_KuCun
        rs.update
        rs.close
        set rs=nothing
        call ok("���ѳɹ��༭������һ����Ʒ��Ϣ��","product_info_list.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ʒ��Ϣ�༭</title>
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
</script>
<script language = "JavaScript">   
var imgObj;
function checkImg(theURL,winName){
  // �����Ƿ��Ѵ���
  if (typeof(imgObj) == "object"){
    // �Ƿ���ȡ����ͼ��ĸ߶ȺͿ��
    if ((imgObj.width != 0) && (imgObj.height != 0))
      // ����ȡ�õ�ͼ��߶ȺͿ�����õ������ڵĸ߶����ȣ����򿪸ô���
      // ���е����� 20 �� 30 �����õĴ��ڱ߿���ͼƬ��ļ����
      OpenFullSizeWindow(theURL,winName, ",width=" + (imgObj.width+20) + ",height=" + (imgObj.height+30));
    else
      // ��Ϊͨ�� Image ����̬װ��ͼƬ�������������õ�ͼƬ�Ŀ�Ⱥ͸߶ȣ�����ÿ��100�����ظ����ü��
      setTimeout("checkImg('" + theURL + "','" + winName + "')", 100)
  }
}

function OpenFullSizeWindow(theURL,winName,features) {
  var aNewWin, sBaseCmd;
  // ����������۲���
  sBaseCmd = "toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,";
  // �����Ƿ����� checkImg 
  if (features == null || features == ""){
    // ����ͼ�����
    imgObj = new Image();
    // ����ͼ��Դ
    imgObj.src = theURL;
    // ��ʼ��ȡͼ���С
    checkImg(theURL, winName)
  }
  else{
    // �򿪴���
    aNewWin = window.open(theURL,winName, sBaseCmd + features);
    // �۽�����
    aNewWin.focus();
  }
}

function loaded(myimg,mywidth,myheight)
{
 var tmp_img = new Image();
 tmp_img.src = myimg.src;
 image_x = tmp_img.width;
 image_y=tmp_img.height;

 if(image_x > mywidth)
 {
  tmp_img.height = image_y * mywidth / image_x;
  tmp_img.width = mywidth;

  if(tmp_img.height > myheight)
  {
   tmp_img.width = tmp_img.width * myheight / tmp_img.height;
   tmp_img.height=myheight;
  }
 }
 else if(image_y > myheight)
 {
  tmp_img.width = image_x * myheight / image_y;
  tmp_img.height=myheight;
  
  if(tmp_img.width > mywidth)
  {
   tmp_img.height = tmp_img.height * mywidth / tmp_img.width;
   tmp_img.width=mywidth;
  }
 }
  
 myimg.width = tmp_img.width;
 myimg.height = tmp_img.height;
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
<form action="Product_Info_Modi.asp" method="post" name="form1">
<input type="hidden" name="action" value="save">
<input type="hidden" name="id" value="<%=id%>"> 
    <tr>
		<td colspan="3" class="title">��Ʒ��Ϣ�༭</td>
	</tr>
	<tr>
		<td>��Ʒ���Ƽ����</td>
		<td colspan="2">
		<input type="text" name="product_info_name" size="30" value="<%=product_info_name%>"></td>
	</tr>
	<tr>
		<td>��Ʒ���ţ�</td>
		<td colspan="2">
		<input type="text" name="product_info_no" size="30" value="<%=product_info_no%>"></td>
	</tr>
	<tr>
		<td>������Ʒ���</td>
		<td colspan="2">
			<select name="bid" onChange="changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)">
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
           		<%if sid<>"" then%><option value="<%=sid%>" selected><%=SClass1%></option><%end if%> 
         	</select>
		</td>
	</tr>
  
	<tr>
		<td>�г��ۣ�</td>
		<td colspan="2">
		<input type="text" name="product_info_PriceM" size="30" value="<%=FormatNumber(product_info_PriceM,2,-1)%>"></td>
	</tr>
	<tr>
		<td>��վ�ۣ�</td>
		<td colspan="2">
		<input type="text" name="product_info_PriceS" size="30" value="<%=FormatNumber(product_info_PriceS,2,-1)%>"></td>
	</tr>
	<tr>
		<td>СͼƬ��</td>
		<td>
		        <input type="text" name="product_info_PicS" size="30" value="<%=product_info_PicS%>">
		        <input type="button" value=">>����ϴ���ƷСͼƬ" name="action0" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicS','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=400,height=100')">
		</td>
		<td rowspan="2" align="center">
		<a target="_blank" title="����鿴��Ʒ��ͼƬ" href="../uploadpic/<%=product_info_PicB%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=product_info_PicS%> border=0 onload='loaded(this,80,80)' ><br>���
		�鿴��һ�Ŵ�ͼ</a></td>
	</tr>
<tr>
		<td>��ͼƬ��</td>
		<td>
		        <input type="text" name="product_info_PicB" size="30" value="<%=product_info_PicB%>">
		        <input type="button" value=">>����ϴ���Ʒ��ͼƬ" name="action1" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=400,height=100')">
        		<br>
				<input type="checkbox" name="MorePic" value="1"  onClick='showlist(paipai);' <%if product_info_PicB2<>"" or product_info_PicB3<>"" then%>checked<%end if%>>Ҫ�ϴ�������Ʒ��ͼƬ,���ڷ����ڴ�(<font color="#808080">��๲֧��������Ʒ��ͼƬ</font>)</td>
	</tr>
	<tr id=paipai <%if product_info_PicB2<>"" or product_info_PicB3<>"" then%><%else%>style="display:none"<%end if%>>
		<td>��ͼ�ϴ���</td>
		<td colspan="2">
			<table border="0" width="100%" id="table1" cellpadding="2" style="border-collapse: collapse">
				<tr>
					<td>�ڶ�����Ʒ��ͼ��<input type="text" name="product_info_PicB2" size="30" readonly value="<%=product_info_PicB2%>"> <a href="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB2','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=410,height=230')">
			<img src="images/upload.gif" alt="�ϴ�ͼƬ" style="cursor: hand;" onMouseOver="window.status='ʹ��ϵͳ�Դ����ϴ������ϴ�ͼƬ';return true;" onMouseOut="window.status='';return true;" border="0"></a></td>
					<td align=center><%if product_info_PicB2<>"" then%><a target="_blank" title="����鿴�ڶ�����Ʒ��ͼƬ" href="../uploadpic/<%=product_info_PicB2%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=product_info_PicB2%> border=0 onload='loaded(this,80,80)' ><br>����鿴�ڶ��Ŵ�ͼ</a><%end if%></td>
				</tr>
				<tr>
					<td>��������Ʒ��ͼ��<input type="text" name="product_info_PicB3" size="30" readonly value="<%=product_info_PicB3%>"> <a href="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB3','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=410,height=230')">
			<img src="images/upload.gif" alt="�ϴ�ͼƬ" style="cursor: hand;" onMouseOver="window.status='ʹ��ϵͳ�Դ����ϴ������ϴ�ͼƬ';return true;" onMouseOut="window.status='';return true;" border="0"></a>
					</td>
					<td align=center><%if product_info_PicB3<>"" then%><a target="_blank" title="����鿴��������Ʒ��ͼƬ" href="../uploadpic/<%=product_info_PicB3%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=product_info_PicB3%> border=0 onload='loaded(this,80,80)' ><br>����鿴�����Ŵ�ͼ</a><%end if%></td>
				</tr>
			</table>
	</tr>
	<tr>
		<td>�������</td>
		<td colspan="2">
		<input type="text" name="product_info_KuCun" size="30" value="<%=product_info_KuCun%>">��</td>
	</tr>
	<tr>
		<td>��ϸ������</td>
		<td colspan="2">
		<textarea cols=80 rows=20 id="content" name="Content"><%= Server.HTMLEncode(product_info_Detail) %></textarea>
		</td>
	</tr>
	<tr>
		<td>��Ʒ���ԣ�</td>
		<td colspan="2"><input type="checkbox" name="product_info_flag" value="1" <%if instr(product_info_flag,1) then response.write "checked" %>>��Ʒ&nbsp; 
		<input type="checkbox" name="product_info_flag" value="2" <%if instr(product_info_flag,2) then response.write "checked" %>>�Ƽ�&nbsp; 
		<input type="checkbox" name="product_info_flag" value="3" <%if instr(product_info_flag,3) then response.write "checked" %>>�ؼ�</td>
	</tr>
	<tr>
		<td>�Ƿ��ϼܣ�</td>
		<td colspan="2"><input type="radio" value="0" name="product_info_OnOff" <%if product_info_OnOff=0 then response.write "checked" %>>�ϼ�(��ʾ)&nbsp;&nbsp;
		    <input type="radio" value="1" name="product_info_OnOff" <%if product_info_OnOff=1 then response.write "checked" %>>�¼�(����) </td>
	</tr>
	<tr>
		<td>��</td>
		<td colspan="2"><input type="submit" value="�ύ" name="Submit1">&nbsp; 
		    <input type="reset" value="����" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>
