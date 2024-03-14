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
  response.write("<script>alert(""参数错误!"");location.href=""product_info_List.asp"";</script>")
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
	    ErrMsg=ErrMsg & "<li>商品ID不能为空！</li>"
    end if
    if product_info_name="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>商品名称不能为空！</li>"
    end if
    if bid="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>商品大类别必须选择！</li>"
    end if
    if sid="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>商品小类别必须选择！</li>"
    end if
    if product_info_PriceS="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>本站价格不能为空！</li>"
    end if

    if product_info_Detail="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>商品详细描述不能为空！</li>"
    end if
                     
    if FoundErr<>True then
        set rs=server.createobject("adodb.recordset")
        sql="select * from product_info Where product_info_name='"&product_info_name&"' and id<>"&id
        rs.open sql,conn,1,1
      	if not rs.eof and rs.bof then
       		response.write "<script language='javascript'>"
        	response.write "alert('出错了，商品标题重复，请重新录入！');"
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
        call ok("您已成功编辑更新了一条商品信息！","product_info_list.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品信息编辑</title>
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
		<td colspan="3" class="title">商品信息编辑</td>
	</tr>
	<tr>
		<td>商品名称及规格：</td>
		<td colspan="2">
		<input type="text" name="product_info_name" size="30" value="<%=product_info_name%>"></td>
	</tr>
	<tr>
		<td>商品货号：</td>
		<td colspan="2">
		<input type="text" name="product_info_no" size="30" value="<%=product_info_no%>"></td>
	</tr>
	<tr>
		<td>所属商品类别：</td>
		<td colspan="2">
			<select name="bid" onChange="changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)">
		    	<option>请选择大类</option>
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
		   		<option value="" <%if sid="" or null(sid) then response.write "selected" %>>请选择小类</option>		  
           		<%if sid<>"" then%><option value="<%=sid%>" selected><%=SClass1%></option><%end if%> 
         	</select>
		</td>
	</tr>
  
	<tr>
		<td>市场价：</td>
		<td colspan="2">
		<input type="text" name="product_info_PriceM" size="30" value="<%=FormatNumber(product_info_PriceM,2,-1)%>"></td>
	</tr>
	<tr>
		<td>本站价：</td>
		<td colspan="2">
		<input type="text" name="product_info_PriceS" size="30" value="<%=FormatNumber(product_info_PriceS,2,-1)%>"></td>
	</tr>
	<tr>
		<td>小图片：</td>
		<td>
		        <input type="text" name="product_info_PicS" size="30" value="<%=product_info_PicS%>">
		        <input type="button" value=">>点此上传商品小图片" name="action0" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicS','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=400,height=100')">
		</td>
		<td rowspan="2" align="center">
		<a target="_blank" title="点击查看商品大图片" href="../uploadpic/<%=product_info_PicB%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=product_info_PicS%> border=0 onload='loaded(this,80,80)' ><br>点击
		查看第一张大图</a></td>
	</tr>
<tr>
		<td>大图片：</td>
		<td>
		        <input type="text" name="product_info_PicB" size="30" value="<%=product_info_PicB%>">
		        <input type="button" value=">>点此上传商品大图片" name="action1" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=400,height=100')">
        		<br>
				<input type="checkbox" name="MorePic" value="1"  onClick='showlist(paipai);' <%if product_info_PicB2<>"" or product_info_PicB3<>"" then%>checked<%end if%>>要上传多张商品大图片,请在方框内打勾(<font color="#808080">最多共支持三张商品大图片</font>)</td>
	</tr>
	<tr id=paipai <%if product_info_PicB2<>"" or product_info_PicB3<>"" then%><%else%>style="display:none"<%end if%>>
		<td>多图上传：</td>
		<td colspan="2">
			<table border="0" width="100%" id="table1" cellpadding="2" style="border-collapse: collapse">
				<tr>
					<td>第二张商品大图：<input type="text" name="product_info_PicB2" size="30" readonly value="<%=product_info_PicB2%>"> <a href="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB2','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=410,height=230')">
			<img src="images/upload.gif" alt="上传图片" style="cursor: hand;" onMouseOver="window.status='使用系统自带的上传程序上传图片';return true;" onMouseOut="window.status='';return true;" border="0"></a></td>
					<td align=center><%if product_info_PicB2<>"" then%><a target="_blank" title="点击查看第二张商品大图片" href="../uploadpic/<%=product_info_PicB2%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=product_info_PicB2%> border=0 onload='loaded(this,80,80)' ><br>点击查看第二张大图</a><%end if%></td>
				</tr>
				<tr>
					<td>第三张商品大图：<input type="text" name="product_info_PicB3" size="30" readonly value="<%=product_info_PicB3%>"> <a href="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB3','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=410,height=230')">
			<img src="images/upload.gif" alt="上传图片" style="cursor: hand;" onMouseOver="window.status='使用系统自带的上传程序上传图片';return true;" onMouseOut="window.status='';return true;" border="0"></a>
					</td>
					<td align=center><%if product_info_PicB3<>"" then%><a target="_blank" title="点击查看第三张商品大图片" href="../uploadpic/<%=product_info_PicB3%>" onClick="OpenFullSizeWindow(this.href,'','');return false"><img src=../uploadpic/<%=product_info_PicB3%> border=0 onload='loaded(this,80,80)' ><br>点击查看第三张大图</a><%end if%></td>
				</tr>
			</table>
	</tr>
	<tr>
		<td>库存量：</td>
		<td colspan="2">
		<input type="text" name="product_info_KuCun" size="30" value="<%=product_info_KuCun%>">件</td>
	</tr>
	<tr>
		<td>详细描述：</td>
		<td colspan="2">
		<textarea cols=80 rows=20 id="content" name="Content"><%= Server.HTMLEncode(product_info_Detail) %></textarea>
		</td>
	</tr>
	<tr>
		<td>商品特性：</td>
		<td colspan="2"><input type="checkbox" name="product_info_flag" value="1" <%if instr(product_info_flag,1) then response.write "checked" %>>新品&nbsp; 
		<input type="checkbox" name="product_info_flag" value="2" <%if instr(product_info_flag,2) then response.write "checked" %>>推荐&nbsp; 
		<input type="checkbox" name="product_info_flag" value="3" <%if instr(product_info_flag,3) then response.write "checked" %>>特价</td>
	</tr>
	<tr>
		<td>是否上架：</td>
		<td colspan="2"><input type="radio" value="0" name="product_info_OnOff" <%if product_info_OnOff=0 then response.write "checked" %>>上架(显示)&nbsp;&nbsp;
		    <input type="radio" value="1" name="product_info_OnOff" <%if product_info_OnOff=1 then response.write "checked" %>>下架(隐藏) </td>
	</tr>
	<tr>
		<td>　</td>
		<td colspan="2"><input type="submit" value="提交" name="Submit1">&nbsp; 
		    <input type="reset" value="重置" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>
