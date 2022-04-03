<!--#include file="admin_check.asp"-->
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="../include/pages.asp"-->
<%
Search_KeyWord	= my_request("KeyWord",0) 					'商品名称关键字
Search_cid		= my_request("cid",0)						'商品类别id
Search_Detail	= my_request("prod_info_detail",0)		'商品内容关键字
Search_PriceSMin= my_request("prod_info_PriceSMin",1)	'本站价格范围小值
Search_PriceSMax= my_request("prod_info_PriceSMax",1)	'本站价格范围大值
Search_Sort		= my_request("sort",1)						'结果排序

Search=""
if Search_KeyWord<>"" then
    Search=Search & " and prod_info_name like '%"&Search_KeyWord&"%'"
end if

if Search_cid<>"" then
    Search=Search & " and cid="&Search_cid
end if

if Search_Detail<>"" then
    Search=Search & " and prodcut_info_Detail like '%"&Search_Detail&"%'"
end if

if Search_PriceSMin<>"" and Search_PriceSMax<>"" then 
    Search=Search & " and (prod_info_PriceS Between "&Search_PriceSMin&" and "&Search_PriceSMax&")"
end if

if Search_PriceSMin<>"" and Search_PriceSMax="" then 
    Search=Search & " and prod_info_PriceS>"&Search_PriceSMin
end if

if Search_PriceSMin="" and Search_PriceSMax<>"" then 
	Search=Search & " and prod_info_PriceS<"&Search_PriceSMax
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
        orderby=" order by prod_info_name"
    case 6
        orderby=" order by prod_info_hitnums desc"
    case else
        orderby=" order by addtime desc"
    end select     
else
    orderby=" order by addtime desc"
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品信息管理</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language = "JavaScript">   
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
select case action
case "删除"
	Call ProdDel()
case "下架"
	Call ProdOff()
case "上架"
	Call ProdOn()
case "设为新品"
	Call ProdNewY()
case "取消新品"
	Call ProdNewN()
case "设为推荐"
	Call ProdTJY()	
case "取消推荐"
	Call ProdTJN()
case "设为特价"
	Call ProdSpecY()	
case "取消特价"
	Call ProdSpecN()
case "确定"
	Call ProdClass()
end select

'过程：批量删除商品
sub ProdDel()
    id=my_request("id",0)
    if id<>"" then
       pp=ubound(split(id,","))+1 '判断数组id中共有几维
       for v=1 to pp
          id=request("id")(v)
          
          sql="select prod_info_PicB,prod_info_PicS from prod_info where id="&id
          set rs=conn.execute (sql)
          prod_info_PicB  =rs("prod_info_PicB")
          prod_info_PicS=rs("prod_info_PicS")
          rs.close
          set rs=nothing
          
          conn.execute ("delete from [prod_info] where id="&id)
          
          //删除相应商品图片
          Dbpath="../uploadpic/"&prod_info_PicS
          Dbpath=server.mappath(Dbpath)
          bkfolder="../uploadpic"
          Set Fso=server.createobject("scripting.filesystemobject")
          if fso.fileexists(dbpath) then
              If CheckDir(bkfolder) = True Then
                  fso.DeleteFile dbpath
              end if
          end if
          Set fso = nothing

          Dbpath1="../uploadpic/"&prod_info_PicB
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
       response.write "location.href='prod_info_list.asp';"
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
    sql="select prod_info_PicB,prod_info_PicS from prod_info where id="&id
    set rs=conn.execute (sql)
    prod_info_PicB=rs("prod_info_PicB")
    prod_info_PicS=rs("prod_info_PicS")
    rs.close
    set rs=nothing

    '判定是否存在小图片文件:
    if fso.FileExists(server.MapPath("uploadpic/"&prod_info_PicS)) then
        '如果存在,删除该文件
        fso.DeleteFile server.MapPath("uploadpic/"&prod_info_PicS),true
        set fso=nothing
    end if
    '判定是否存在大图片文件:
    if fso.FileExists(server.MapPath("uploadpic/"&prod_info_PicB)) then
        '如果存在,删除该文件
        fso.DeleteFile server.MapPath("uploadpic/"&prod_info_PicB),true
        set fso=nothing
        call ok("所选信息已成功删除！","prod_info_list.asp")

    end if
end function

sub ProdOff()
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
            conn.execute ("update prod_info set prod_info_onoff=1 where id="&id)
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已设置为下架！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub

sub ProdOn()
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
            conn.execute ("update prod_info set prod_info_onoff=0 where id="&id)
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已设置为上架！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub

'批量修改分类
sub ProdClass()  
    id=my_request("id",0)
    cid1=my_request("cid1",1)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
            conn.execute ("update prod_info set cid="&cid1&" where id="&id)
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已移动到指定类别！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub

'批量新品
sub ProdNewY()  
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
    		Set rs= Server.CreateObject("ADODB.Recordset")
    		sql="select prod_info_flag from prod_info where id="&id
    		rs.open sql,conn,1,3
    		prod_info_flag1=rs(0)
    		rs.close
    		set rs=nothing
    
   			prod_info_flag1=prod_info_flag1&",1"	
    		conn.execute ("update prod_info set prod_info_flag='"&prod_info_flag1&"' where id="&id)	
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已设置为新品！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub

'批量非新品
sub ProdNewN()  
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
    		Set rs= Server.CreateObject("ADODB.Recordset")
    		sql="select prod_info_flag from prod_info where id="&id
    		rs.open sql,conn,1,3
    		prod_info_flag1=rs(0)
    		rs.close
    		set rs=nothing
    		prod_info_flag1=Replace(prod_info_flag1,"1","")	
    		conn.execute ("update prod_info set prod_info_flag='"&prod_info_flag1&"' where id="&id)	
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已取消新品！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub

'批量新品
sub ProdTjY()  
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
    		Set rs= Server.CreateObject("ADODB.Recordset")
    		sql="select prod_info_flag from prod_info where id="&id
    		rs.open sql,conn,1,3
    		prod_info_flag1=rs(0)
    		rs.close
    		set rs=nothing
    
   			prod_info_flag1=prod_info_flag1&",2"	
    		conn.execute ("update prod_info set prod_info_flag='"&prod_info_flag1&"' where id="&id)	
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已设置为推荐！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub

'批量非推荐
sub ProdTjN()  
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
    		Set rs= Server.CreateObject("ADODB.Recordset")
    		sql="select prod_info_flag from prod_info where id="&id
    		rs.open sql,conn,1,3
    		prod_info_flag1=rs(0)
    		rs.close
    		set rs=nothing
    		prod_info_flag1=Replace(prod_info_flag1,"2","")	
    		conn.execute ("update prod_info set prod_info_flag='"&prod_info_flag1&"' where id="&id)	
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已取消推荐！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub

'批量特价
sub ProdSpecY()  
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
    		Set rs= Server.CreateObject("ADODB.Recordset")
    		sql="select prod_info_flag from prod_info where id="&id
    		rs.open sql,conn,1,3
    		prod_info_flag1=rs(0)
    		rs.close
    		set rs=nothing
    
   			prod_info_flag1=prod_info_flag1&",3"	
    		conn.execute ("update prod_info set prod_info_flag='"&prod_info_flag1&"' where id="&id)	
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已设置为特价！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub

'批量非特价
sub ProdSpecN()  
    id=my_request("id",0)
    if id<>"" then
        pp=ubound(split(id,","))+1 '判断数组id中共有几维
        for v=1 to pp
            id=request("id")(v)
    		Set rs= Server.CreateObject("ADODB.Recordset")
    		sql="select prod_info_flag from prod_info where id="&id
    		rs.open sql,conn,1,3
    		prod_info_flag1=rs(0)
    		rs.close
    		set rs=nothing
    		prod_info_flag1=Replace(prod_info_flag1,"3","")	
    		conn.execute ("update prod_info set prod_info_flag='"&prod_info_flag1&"' where id="&id)	
        next
       	response.write "<script language='javascript'>"
       	response.write "alert('所选商品均已取消特价！');"
        response.write "location.href='prod_info_list.asp';"
        response.write "</script>"
    end if
end sub
%>
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td class="title" colspan="11">商品信息管理</td>
	</tr>
	<form name=form11 action=prod_info_list.asp method=get>
	<tr>
		<td class="altbg2" colspan="11">
		<p align="center">
				<img border="0" src="images/write11.gif" align="middle"><a href="Prod_Info_Add.asp"><font color="#FF6600">添加商品</font></a><font color="#FF6600">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</font><b>商品搜索：</b><input type="text" name="KeyWord" size="20">
		<input type="submit" value="搜索" name="B1">&nbsp; 
		<a href="prod_Info_Search.asp">高级搜索</a>&nbsp;&nbsp; <b>
		按类别筛选：</b><select onchange="location=this.options[this.selectedIndex].value;" size="1" name="qs">
		     <option value=prod_info_list.asp>显示所有商品</option>
		    <%
		     sql="select cid,prod_class_name from prod_class order by cid desc"
		     set rs=conn.execute (sql)
		     set cid1=rs(0)
		     set prod_class_name1=rs(1)
		     do while not rs.eof
		    %>
		    <option value="prod_info_list.asp?cid=<%=cid1%>" <%if cid1=cint(cid) then response.write "Selected" %>><%=prod_class_name1%></option>
		    <%
		     rs.movenext
		     loop
		     rs.close
		     set rs=nothing
		    %>
		 </select></td>
	</tr>
	</form>
	<tr>
		<td class="altbg1">选中</td>
		<td class="altbg1">
		<p align="center">商品缩图</td>
		<td class="altbg1">商品名称</td>
		<td class="altbg1">所属类别</td>
		<td class="altbg1">市场价</td>
		<td class="altbg1">本站价</td>
		<td class="altbg1">特性</td>
		<td class="altbg1">发布时间</td>
		<td class="altbg1">浏览</td>
		<td class="altbg1">
		<p align="center">状态</td>
		<td class="altbg1">
		<p align="center">编辑</td>
	</tr>
	<form name="form1" action="prod_info_List.asp" method="post">
	<%
    set rs=server.createobject("adodb.recordset")
    if Search_KeyWord="" and Search_cid="" and Search_PriceSMin="" and Search_PriceSMax="" and Search_Detail="" and Search_sort="" then
        sql="select id,cid,prod_info_name,prod_info_flag,prod_info_PriceM,prod_info_PriceS,prod_info_PicB,prod_info_PicS,prod_info_hitnums,addtime,prod_info_OnOff,prod_info_AdWord from prod_info order by id desc"
    else
       sql="select id,cid,prod_info_name,prod_info_flag,prod_info_PriceM,prod_info_PriceS,prod_info_PicB,prod_info_PicS,prod_info_hitnums,addtime,prod_info_OnOff,prod_info_AdWord from prod_info where 1=1 "& Search
    end if
    rs.open sql,conn,1,1
    if rs.eof then 
        response.write "<tr><td colspan=11 align=center>目前暂无商品信息,<a href=prod_info_add.asp>请添加新商品信息!</a></td></tr>"
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
        set cid2				=rs(1)
      	set prod_info_name	=rs(2)
      	set prod_info_flag	=rs(3)
      	set prod_info_PriceM	=rs(4)
      	set prod_info_PriceS	=rs(5)
      	set prod_info_PicB	=rs(6)
      	set prod_info_PicS	=rs(7)
      	set prod_info_hitnums=rs(8)
      	set addtime		=rs(9)
      	set prod_info_OnOff  =rs(10)
      	set prod_info_AdWord	=rs(11)


      	
        while not rs.eof and i<=rs.pagesize

      	prod_info_addtime=datevalue(addtime)
      	if prod_info_OnOff=0 then txt_OnOff="<font color=#009933>上架</font>" else txt_OnOff="<font color=#999>下架</font>"
        
        //调出商品类别名称
		sql1="select prod_class_name from prod_Class where cid="&cid2
		set rs1=conn.execute (sql1)
		prod_class_name2=rs1(0)
		rs1.close
		set rs1=nothing
		
        txt=""
        if instr(prod_info_flag,1) then txt="新、"
        if instr(prod_info_flag,2) then txt=txt&"荐、"
        if instr(prod_info_flag,3) then txt=txt&"特"
    %>
   	<tr>
		<td><input type="checkbox" name="id" value="<%=id%>"></td>
		<td>
		<p align="center"><a target="_blank" title="点击查看商品大图片" href="../uploadpic/<%=prod_info_PicB%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=prod_info_PicS%> border=0 onload='loaded(this,80,80)' ></a></td>
		<td><a href=../prod_detail.asp?id=<%=id%> target=_blank><%=prod_info_name%><br><b><font color=#FF0000><%=prod_info_AdWord%></font></b></a></td>
		<td><%=prod_class_name2%></td>
		<td><font color="#C0C0C0">￥<%=FormatNumber(prod_info_PriceM,2,-1)%></font></td>
		<td><b><font color="#FF0000">￥<%=FormatNumber(prod_info_PriceS,2,-1)%></font></b></td>
		<td><%=txt%></td>
		<td><%=prod_info_addtime%></td>
		<td><%=prod_info_HitNums%>次</td>
		<td><%=txt_OnOff%></td>
		<td>
		<p align="center"><a href=prod_info_Modi.asp?id=<%=id%>>编辑</a></td>

	</tr>
	<%
        rs.movenext
        i=i+1
        wend
	%>
	<tr>
		<td colspan="11">
		<table border="0" width="100%" cellpadding="2" cellspacing="1">
			<tr>
				<td colspan="2">
				<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>全选</td>
			</tr>
			<tr>
				<td bgcolor="#FFD1B3" colspan="2">批量操作</td>
			</tr>
			<tr>
				<td bgcolor="#FFDFCA">第一步：<br>
				打勾选中操作项</td>
				<td bgcolor="#FFDFCA">第二步：<br>
				批量操作区</td>
			</tr>
			<tr>
				<td bgcolor="#F3F3F3">　</td>
				<td bgcolor="#F3F3F3">
				<table border="0" width="100%" style="border-collapse: collapse" cellpadding="2">
					<tr>
						<td style="border-bottom: 1px dotted #DDDDDD">批量删除：</td>
						<td style="border-bottom: 1px dotted #DDDDDD">
						<input type="submit" name="action" value="删除" onclick="{if(confirm('删除后将无法恢复，您确定要删除选定的信息吗？')){this.document.form1.submit();return true;}return false;}"></td>
					</tr>
					<tr>
						<td style="border-bottom: 1px dotted #DDDDDD">批量设置上下架：</td>
						<td style="border-bottom: 1px dotted #DDDDDD">
							<input type="submit" name="action" value="下架">&nbsp; 
        					<input type="submit" name="action" value="上架">
						</td>
					</tr>
					<tr>
						<td style="border-bottom: 1px dotted #DDDDDD">批量设置特性：</td>
						<td style="border-bottom: 1px dotted #DDDDDD"> 
        					<input type="submit" name="action" value="设为新品">&nbsp; 
        					<input type="submit" name="action" value="取消新品">&nbsp; 
        					<input type="submit" name="action" value="设为推荐">&nbsp; 
        					<input type="submit" name="action" value="取消推荐">&nbsp;
        					<input type="submit" name="action" value="设为特价">&nbsp; 
        					<input type="submit" name="action" value="取消特价">
						</td>
					</tr>
					<tr>
						<td style="border-bottom: 1px dotted #DDDDDD">批量设置类别：</td>
						<td>移动商品到类别：<select size="1" name="cid1">
		    <%
		     sql="select cid,prod_class_name from prod_class order by cid desc"
		     set rs=conn.execute (sql)
		     set cid2=rs(0)
		     set prod_class_name2=rs(1)
		     do while not rs.eof
		    %>
		    <option value="<%=cid2%>"><%=prod_class_name2%></option>
		    <%
		     rs.movenext
		     loop
		     rs.close
		     set rs=nothing
		    %>
		 </select> <input type="submit" name="action" value="确定">
						</td>
					</tr>
				</table>
				</td>
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

