<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=1
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/pages.asp"-->
<%
Search_KeyWord	= my_request("KeyWord",0) 					'��Ʒ���ƹؼ���
Search_bid		= my_request("bid",1)						'��Ʒ�����id
Search_sid		= my_request("sid",1)						'��ƷС���id
Search_Detail	= my_request("product_info_detail",0)		'��Ʒ���ݹؼ���
Search_PriceSMin= my_request("product_info_PriceSMin",1)	'��վ�۸�ΧСֵ
Search_PriceSMax= my_request("product_info_PriceSMax",1)	'��վ�۸�Χ��ֵ
Search_Sort		= my_request("sort",1)						'�������

Search=""
if Search_KeyWord<>"" then
    Search=Search & " and product_info_name like '%"&Search_KeyWord&"%'"
end if

if Search_bid<>"" then
    Search=Search & " and bid="&Search_bid
end if

if Search_sid<>"" then
    Search=Search & " and sid="&Search_sid
end if

if Search_Detail<>"" then
    Search=Search & " and prodcut_info_Detail like '%"&Search_Detail&"%'"
end if

if Search_PriceSMin<>"" and Search_PriceSMax<>"" then 
    Search=Search & " and (product_info_PriceS Between "&Search_PriceSMin&" and "&Search_PriceSMax&")"
end if

if Search_PriceSMin<>"" and Search_PriceSMax="" then 
    Search=Search & " and product_info_PriceS>"&Search_PriceSMin
end if

if Search_PriceSMin="" and Search_PriceSMax<>"" then 
	Search=Search & " and product_info_PriceS<"&Search_PriceSMax
end if

if Search_Sort<>"" then
    select case Search_Sort
    case 1
        orderby=" order by addtime desc"
    case 2
        orderby=" order by addtime asc"
    case 3
        orderby=" order by id desc"
    case 4 
        orderby=" order by id asc"
    case 5
        orderby=" order by product_info_name"
    case 6
        orderby=" order by product_info_hitnums desc"
    case else
        orderby=" order by addtime desc"
    end select     
else
    orderby=" order by addtime desc"
end if

x=my_request("x",0)
select case x
	case "a1"
   		call a1()
	case "a2"
   		call a2()
	case "a3"
   		call a3()
	case "a4"
   		call a4()
	case "a5"
   		call a5()
	case "a6"
   		call a6()
end select

sub a1()
	id=request("id")
	Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select product_info_flag from product_info where id="&id
    rs.open sql,conn,1,3
    product_info_flag1=rs(0)
    rs.close
    set rs=nothing
    
    product_info_flag1=Replace(product_info_flag1,"1","")	
    conn.execute ("update product_info set product_info_flag='"&product_info_flag1&"' where id="&id)	
end sub

sub a2()
	id=request("id")
	Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select product_info_flag from product_info where id="&id
    rs.open sql,conn,1,3
    product_info_flag1=rs(0)
    rs.close
    set rs=nothing
    
    product_info_flag1=product_info_flag1&",1"	
    conn.execute ("update product_info set product_info_flag='"&product_info_flag1&"' where id="&id)	
end sub

sub a3()
	id=request("id")
	Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select product_info_flag from product_info where id="&id
    rs.open sql,conn,1,3
    product_info_flag1=rs(0)
    rs.close
    set rs=nothing
    
    product_info_flag1=Replace(product_info_flag1,"2","")	
    'response.write product_info_flag1
	'response.end
    conn.execute ("update product_info set product_info_flag='"&product_info_flag1&"' where id="&id)	
end sub

sub a4()
	id=request("id")
	Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select product_info_flag from product_info where id="&id
    rs.open sql,conn,1,3
    product_info_flag1=rs(0)
    rs.close
    set rs=nothing
    
    product_info_flag1=product_info_flag1&",2"	
    conn.execute ("update product_info set product_info_flag='"&product_info_flag1&"' where id="&id)	
end sub

sub a5()
	id=request("id")
	Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select product_info_flag from product_info where id="&id
    rs.open sql,conn,1,3
    product_info_flag1=rs(0)
    rs.close
    set rs=nothing
    
    product_info_flag1=Replace(product_info_flag1,"3","")	
    'response.write product_info_flag1
	'response.end
    conn.execute ("update product_info set product_info_flag='"&product_info_flag1&"' where id="&id)	
end sub

sub a6()
	id=request("id")
	Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select product_info_flag from product_info where id="&id
    rs.open sql,conn,1,3
    product_info_flag1=rs(0)
    rs.close
    set rs=nothing
    
    product_info_flag1=product_info_flag1&",3"	
    conn.execute ("update product_info set product_info_flag='"&product_info_flag1&"' where id="&id)	
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ʒ��Ϣ����</title>
<link rel="stylesheet" type="text/css" href="style.css">
<%
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from prod_smallclass order by prod_smallclass_bid desc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
subcat[0] = new Array("�˴���������С��","<%= trim(rs("prod_smallclass_bid"))%>","");
        <%
        count = 1
        do while not rs.eof 
        ss=trim(rs("prod_smallclass_bid"))
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("prod_smallclass_name"))%>","<%= trim(rs("prod_smallclass_bid"))%>","<%= trim(rs("prod_smallclass_id"))%>");
        <%
        count = count + 1
        rs.movenext
        if trim(rs("prod_smallclass_bid"))<>ss then
        %>
subcat[<%=count%>] = new Array("�˴���������С��","<%= trim(rs("prod_smallclass_bid"))%>","");   
        <%
        count = count + 1
        end if
        loop
        rs.close
        set rs=nothing
        %>
onecount=<%=count%>;

//����л�
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
    
var imgObj;
function checkImg(theURL,winName){
  // �����Ƿ��Ѵ���
  if (typeof(imgObj) == "object"){
    // �Ƿ���ȡ����ͼ��ĸ߶ȺͿ���
    if ((imgObj.width != 0) && (imgObj.height != 0))
      // ����ȡ�õ�ͼ��߶ȺͿ������õ������ڵĸ߶�����ȣ����򿪸ô���
      // ���е����� 20 �� 30 �����õĴ��ڱ߿���ͼƬ��ļ����
      OpenFullSizeWindow(theURL,winName, ",width=" + (imgObj.width+20) + ",height=" + (imgObj.height+30));
    else
      // ��Ϊͨ�� Image ����̬װ��ͼƬ�������������õ�ͼƬ�Ŀ��Ⱥ͸߶ȣ�����ÿ��100�����ظ����ü��
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

//ȫѡ����    
function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }
</script>
<%
action=my_request("action",0)
if action="ɾ��" then
    call proddel()
end if

'���̣�����ɾ����Ʒ
sub proddel()
    id=my_request("id",0)
    if id<>"" then
       pp=ubound(split(id,","))+1 '�ж�����id�й��м�ά
       for v=1 to pp
          id=request("id")(v)
          
          sql="select product_info_PicB,product_info_PicS from product_info where id="&id
          set rs=conn.execute (sql)
          product_info_PicB  =rs("product_info_PicB")
          product_info_PicS=rs("product_info_PicS")
          rs.close
          set rs=nothing
          
          conn.execute ("delete from [product_info] where id="&id)
          
          //ɾ����Ӧ��ƷͼƬ
          Dbpath="../uploadpic/"&product_info_PicS
          Dbpath=server.mappath(Dbpath)
          bkfolder="../uploadpic"
          Set Fso=server.createobject("scripting.filesystemobject")
          if fso.fileexists(dbpath) then
              If CheckDir(bkfolder) = True Then
                  fso.DeleteFile dbpath
              end if
          end if
          Set fso = nothing

          Dbpath1="../uploadpic/"&product_info_PicB
          Dbpath1=server.mappath(Dbpath1)
          bkfolder1="../uploadpic"
          Set Fso=server.createobject("scripting.filesystemobject")
          if fso.fileexists(dbpath1) then
              If CheckDir(bkfolder1) = True Then
                  fso.DeleteFile dbpath1
              end if
          end if
          Set fso = nothing

       next

       response.write "<script language='javascript'>"
       response.write "alert('��ѡ��Ʒ�Ѿ���ɾ����');"
       response.write "location.href='"&url&"';"
       response.write "</script>"
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

function prodimgdel(id)
    set rs=server.CreateObject("adodb.recordset") '���д���������rsΪ��¼��
    Set fso = Server.CreateObject("Scripting.FileSystemObject") '����fso����
    '�жϷ������Ƿ�֧��fos����
    'if err then 
        'err.clear
        'response.Write("���ܽ���fso������ȷ����Ŀռ�֧��fso:��")
        'response.end
    'end if
    
    //������Ʒ��СͼƬ��ַ
    sql="select product_info_PicB,product_info_PicS from product_info where id="&id
    set rs=conn.execute (sql)
    product_info_PicB=rs("product_info_PicB")
    product_info_PicS=rs("product_info_PicS")
    rs.close
    set rs=nothing

    '�ж��Ƿ����СͼƬ�ļ�:
    if fso.FileExists(server.MapPath("uploadpic/"&product_info_PicS)) then
        '�������,ɾ�����ļ�
        fso.DeleteFile server.MapPath("uploadpic/"&product_info_PicS),true
        set fso=nothing
    end if
    '�ж��Ƿ���ڴ�ͼƬ�ļ�:
    if fso.FileExists(server.MapPath("uploadpic/"&product_info_PicB)) then
        '�������,ɾ�����ļ�
        fso.DeleteFile server.MapPath("uploadpic/"&product_info_PicB),true
        set fso=nothing
        call ok("��ѡ��Ϣ�ѳɹ�ɾ����","product_info_list.asp")

    end if
end function
%>
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td class="title" colspan="13">��Ʒ��Ϣ����</td>
	</tr>
	<tr>
		<td class="altbg2" colspan="13">
		<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
			<tr>
				<td align="center">
				<form name=form11 action=product_info_list.asp method=get>
					<b>��Ʒ������</b>
					<input type="text" name="KeyWord" size="30">
					<input type="submit" value=" �� �� " name="B1">&nbsp; 
					<a href="Product_Info_Search.asp">�߼�����</a>
				</form>
				</td>
				<td> 
				<form name="form1" action="Product_info_List.asp" method="get">
					<b>�����ɸѡ��</b>
					<select name="bid" onChange="changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)">
		    		<option value="">��ѡ�����</option>
		    		<%
		    		sql="select * from prod_bigclass order by prod_bigclass_id desc"
		    		set rs=conn.execute (sql)
		    		do while not rs.eof
		    		%>
		    		<option value="<%=rs("prod_bigclass_id")%>"><%=rs("prod_bigclass_name")%></option>
		    		<%
		    		rs.movenext
		    		loop
		    		rs.close
		    		set rs=nothing
		    		%>
            		</select>&nbsp; 
            		<select name="sid"> 
            		<option value="">��ѡ��С��</option>
            		</select>
            		<input type="submit" value="�ύ">
				</form>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	</form>
	<tr>
		<td class="altbg1">ѡ��</td>
		<td class="altbg1">
		<p align="center">��Ʒ��ͼ</td>
		<td class="altbg1">��Ʒ����</td>
		<td class="altbg1">�������</td>
		<td class="altbg1">�г���</td>
		<td class="altbg1">��վ��</td>
		<td class="altbg1">��Ʒ</td>
		<td class="altbg1">�Ƽ�</td>
		<td class="altbg1">�ؼ�</td>
		<td class="altbg1">����ʱ��</td>
		<td class="altbg1">���</td>
		<td class="altbg1">
		<p align="center">״̬</td>
		<td class="altbg1">
		<p align="center">�༭</td>
	</tr>
	<form name="form2" action="product_info_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    if Search_KeyWord="" and Search_bid="" and Search_sid="" and Search_PriceSMin="" and Search_PriceSMax="" and Search_Detail="" and Search_sort="" then
        sql="select id,bid,sid,product_info_name,product_info_flag,product_info_PriceM,product_info_PriceS,product_info_PicB,product_info_PicS,product_info_hitnums,addtime,product_info_OnOff from product_info order by id desc"
    else
        sql="select id,bid,sid,product_info_name,product_info_flag,product_info_PriceM,product_info_PriceS,product_info_PicB,product_info_PicS,product_info_hitnums,addtime,product_info_OnOff from product_info where 1=1 "& Search
    end if
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=13 align=center>Ŀǰ������Ʒ��Ϣ,<a href=product_info_add.asp>����������Ʒ��Ϣ!</a></td></tr>"
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
        
        set id					=rs(0)
        set bid2				=rs(1)
        set sid2				=rs(2)
      	set product_info_name	=rs(3)
      	set product_info_flag	=rs(4)
      	set product_info_PriceM	=rs(5)
      	set product_info_PriceS	=rs(6)
      	set product_info_PicB	=rs(7)
      	set product_info_PicS	=rs(8)
      	set product_info_hitnums=rs(9)
      	set addtime				=rs(10)
      	set product_info_OnOff  =rs(11)
      	set product_info_AdWord	=rs(12)

      	product_info_addtime=datevalue(addtime)
      	if product_info_OnOff=0 then txt_OnOff="<font color=#0000FF>��</font>" else txt_OnOff="<font color=#FF0000>��</font>"
      	
        while not rs.eof and i<=rs.pagesize
        
        if len(product_info_name)>18 then set product_info_name=left(product_info_name,16)&"...."

        '������Ʒ�������
		sql1="select prod_BigClass_name from prod_BigClass where prod_BigClass_id="&Bid2
		set rs1=conn.execute (sql1)
		BClass1=rs1("prod_BigClass_name")
  		rs1.close
  		set rs1=nothing
  		
		sql2="select prod_SmallClass_name from prod_SmallClass where prod_SmallClass_id="&Sid2
  		set rs2=conn.execute (sql2)
  		SClass1=rs2("prod_SmallClass_name")
  		rs2.close
  		set rs2=nothing

        txt=""
        if instr(product_info_flag,1) then 
        	txt1="<a href=?x=a1&id="&id&"><b><font color=#0000FF>��</font></b></a>" 
        else
        	txt1="<a href=?x=a2&id="&id&"><b><font color=#FF3300>��</font></b></a>" 
        end if
        
        if instr(product_info_flag,2) then 
        	txt2="<a href=?x=a3&id="&id&"><b><font color=#0000FF>��</font></b></a>"
        else
        	txt2="<a href=?x=a4&id="&id&"><b><font color=#FF3300>��</font></b></a>"
        end if
        
        if instr(product_info_flag,3) then 
        	txt3="<a href=?x=a5&id="&id&"><b><font color=#0000FF>��</font></b></a>"
        else
        	txt3="<a href=?x=a6&id="&id&"><b><font color=#FF3300>��</font></b></a>"
		end if
    %>
   	<tr>
		<td><input type="checkbox" name="id" value="<%=id%>"></td>
		<td>
		<p align="center"><a target="_blank" title="����鿴��Ʒ��ͼƬ" href="../uploadpic/<%=product_info_PicB%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=product_info_PicS%> border=0 onload='loaded(this,80,80)' ></a></td>
		<td><a href=product_info_Modi.asp?id=<%=id%>><%=product_info_name%><br><b><font color=#FF0000><%=product_info_AdWord%></font></b></a></td>
		<td><%=BClass1%> &raquo; <%=SClass1%></td>
		<td><font color="#C0C0C0"><%=FormatNumber(product_info_PriceM,2,-1)%></font></td>
		<td><b><font color="#FF6600"><%=FormatNumber(product_info_PriceS,2,-1)%></font></b></td>
		<td align="center"><%=txt1%></td>
		<td align="center"><%=txt2%></td>
		<td align="center"><%=txt3%></td>
		<td><%=product_info_addtime%></td>
		<td><%=product_info_HitNums%></td>
		<td align=center><%=txt_OnOff%></td>
		<td align="center"><a href=product_info_Modi.asp?id=<%=id%>><img src=images/edititem.gif border=0></a></td>

	</tr>
	<%
        rs.movenext
        i=i+1
        wend
	%>
	<tr>
		<td colspan="13">
		<table border="0" width="100%" id="table2">
			<tr>
				<td>
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
        <input type="submit" name="action" value="ɾ��" onclick="{if(confirm('ɾ�����޷��ָ�����ȷ��Ҫɾ��ѡ������Ϣ��')){this.document.form1.submit();return true;}return false;}">&nbsp;
		<input type="button" value="������Ʒ��Ϣ" name="action1" onclick="window.location='product_info_add.asp'"></td>
				<td>
				<p align="right"><font face="����">��</font>˵��<font face="����">��</font>����<font color="#0000FF">��</font>����ʾ�ϼ���Ʒ����<font color="#FF0000">��</font>����ʾ�¼���Ʒ��</td>
			</tr>
		</table>
		</td>
	</tr>
    <input type=hidden name=pagenow value=<%=page%>>
    <%
        call PageControl(iCount,maxpage,page)
    end if
    rs.close
    set rs=nothing
    %>
</form>
</tbody>
</table>

</body>

</html>