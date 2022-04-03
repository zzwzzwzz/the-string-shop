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
Search_KeyWord	= my_request("KeyWord",0) 					'商品名称关键字
Search_bid		= my_request("bid",1)						'商品大类别id
Search_sid		= my_request("sid",1)						'商品小类别id
Search_Detail	= my_request("product_info_detail",0)		'商品内容关键字
Search_PriceSMin= my_request("product_info_PriceSMin",1)	'本站价格范围小值
Search_PriceSMax= my_request("product_info_PriceSMax",1)	'本站价格范围大值
Search_Sort		= my_request("sort",1)						'结果排序

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
<title>商品信息管理</title>
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
subcat[0] = new Array("此大类下所有小类","<%= trim(rs("prod_smallclass_bid"))%>","");
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
subcat[<%=count%>] = new Array("此大类下所有小类","<%= trim(rs("prod_smallclass_bid"))%>","");   
        <%
        count = count + 1
        end if
        loop
        rs.close
        set rs=nothing
        %>
onecount=<%=count%>;

//类别切换
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
  // 对象是否已创建
  if (typeof(imgObj) == "object"){
    // 是否已取得了图像的高度和宽度
    if ((imgObj.width != 0) && (imgObj.height != 0))
      // 根据取得的图像高度和宽度设置弹出窗口的高度与宽度，并打开该窗口
      // 其中的增量 20 和 30 是设置的窗口边框与图片间的间隔量
      OpenFullSizeWindow(theURL,winName, ",width=" + (imgObj.width+20) + ",height=" + (imgObj.height+30));
    else
      // 因为通过 Image 对象动态装载图片，不可能立即得到图片的宽度和高度，所以每隔100毫秒重复调用检查
      setTimeout("checkImg('" + theURL + "','" + winName + "')", 100)
  }
}

function OpenFullSizeWindow(theURL,winName,features) {
  var aNewWin, sBaseCmd;
  // 弹出窗口外观参数
  sBaseCmd = "toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,";
  // 调用是否来自 checkImg 
  if (features == null || features == ""){
    // 创建图像对象
    imgObj = new Image();
    // 设置图像源
    imgObj.src = theURL;
    // 开始获取图像大小
    checkImg(theURL, winName)
  }
  else{
    // 打开窗口
    aNewWin = window.open(theURL,winName, sBaseCmd + features);
    // 聚焦窗口
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
    call proddel()
end if

//过程：批量删除商品
sub proddel()
    id=my_request("id",0)
    if id<>"" then
       pp=ubound(split(id,","))+1 '判断数组id中共有几维
       for v=1 to pp
          id=request("id")(v)
          
          sql="select product_info_PicB,product_info_PicS from product_info where id="&id
          set rs=conn.execute (sql)
          product_info_PicB  =rs("product_info_PicB")
          product_info_PicS=rs("product_info_PicS")
          rs.close
          set rs=nothing
          
          conn.execute ("delete from [product_info] where id="&id)
          
          //删除相应商品图片
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
       response.write "alert('所选商品已经被删除！');"
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
    set rs=server.CreateObject("adodb.recordset") '该行代码是设置rs为记录集
    Set fso = Server.CreateObject("Scripting.FileSystemObject") '建立fso对象
    '判断服务器是否支持fos对象
    'if err then 
        'err.clear
        'response.Write("不能建立fso对象，请确保你的空间支持fso:！")
        'response.end
    'end if
    
    //调出商品大小图片地址
    sql="select product_info_PicB,product_info_PicS from product_info where id="&id
    set rs=conn.execute (sql)
    product_info_PicB=rs("product_info_PicB")
    product_info_PicS=rs("product_info_PicS")
    rs.close
    set rs=nothing

    '判定是否存在小图片文件:
    if fso.FileExists(server.MapPath("uploadpic/"&product_info_PicS)) then
        '如果存在,删除该文件
        fso.DeleteFile server.MapPath("uploadpic/"&product_info_PicS),true
        set fso=nothing
    end if
    '判定是否存在大图片文件:
    if fso.FileExists(server.MapPath("uploadpic/"&product_info_PicB)) then
        '如果存在,删除该文件
        fso.DeleteFile server.MapPath("uploadpic/"&product_info_PicB),true
        set fso=nothing
        call ok("所选信息已成功删除！","product_info_list.asp")

    end if
end function
%>
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td class="title" colspan="13">商品信息管理</td>
	</tr>
	<tr>
		<td class="altbg2" colspan="13">
		<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
			<tr>
				<td align="center">
				<form name=form11 action=product_info_list.asp method=get>
					<b>商品搜索：</b>
					<input type="text" name="KeyWord" size="30">
					<input type="submit" value=" 搜 索 " name="B1">&nbsp; 
					<a href="Product_Info_Search.asp">高级搜索</a>
				</form>
				</td>
				<td> 
				<form name="form1" action="Product_info_List.asp" method="get">
					<b>按类别筛选：</b>
					<select name="bid" onChange="changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)">
		    		<option value="">请选择大类</option>
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
            		<option value="">请选择小类</option>
            		</select>
            		<input type="submit" value="提交">
				</form>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	</form>
	<tr>
		<td class="altbg1">选</td>
		<td class="altbg1">
		<p align="center">商品缩图</td>
		<td class="altbg1">商品名称</td>
		<td class="altbg1">所属类别</td>
		<td class="altbg1">市场价</td>
		<td class="altbg1">本站价</td>
		<td class="altbg1">新品</td>
		<td class="altbg1">推荐</td>
		<td class="altbg1">特价</td>
		<td class="altbg1">发布时间</td>
		<td class="altbg1">浏览</td>
		<td class="altbg1">
		<p align="center">状态</td>
		<td class="altbg1">
		<p align="center">编辑</td>
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
        response.write "<tr><td colspan=13 align=center>目前暂无商品信息,<a href=product_info_add.asp>请添加新商品信息!</a></td></tr>"
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
      	if product_info_OnOff=0 then txt_OnOff="<font color=#0000FF>↑</font>" else txt_OnOff="<font color=#FF0000>↓</font>"
      	
        while not rs.eof and i<=rs.pagesize
        
        if len(product_info_name)>18 then set product_info_name=left(product_info_name,16)&"...."

        //调出商品类别名称
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
        	txt1="<a href=?x=a1&id="&id&"><b><font color=#0000FF>√</font></b></a>" 
        else
        	txt1="<a href=?x=a2&id="&id&"><b><font color=#FF3300>×</font></b></a>" 
        end if
        
        if instr(product_info_flag,2) then 
        	txt2="<a href=?x=a3&id="&id&"><b><font color=#0000FF>√</font></b></a>"
        else
        	txt2="<a href=?x=a4&id="&id&"><b><font color=#FF3300>×</font></b></a>"
        end if
        
        if instr(product_info_flag,3) then 
        	txt3="<a href=?x=a5&id="&id&"><b><font color=#0000FF>√</font></b></a>"
        else
        	txt3="<a href=?x=a6&id="&id&"><b><font color=#FF3300>×</font></b></a>"
		end if
    %>
   	<tr>
		<td><input type="checkbox" name="id" value="<%=id%>"></td>
		<td>
		<p align="center"><a target="_blank" title="点击查看商品大图片" href="../uploadpic/<%=product_info_PicB%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=product_info_PicS%> border=0 onload='loaded(this,80,80)' ></a></td>
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
		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选 
        <input type="submit" name="action" value="删除" onclick="{if(confirm('删除后将无法恢复，您确定要删除选定的信息吗？')){this.document.form1.submit();return true;}return false;}">&nbsp;
		<input type="button" value="添加商品信息" name="action1" onclick="window.location='product_info_add.asp'"></td>
				<td>
				<p align="right"><font face="宋体">【</font>说明<font face="宋体">】</font>：“<font color="#0000FF">↑</font>”表示上架商品，“<font color="#FF0000">↓</font>”表示下架商品。</td>
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