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
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from product_info Where product_info_name='"&product_info_name&"'"
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
        call ok("您已成功添加了一条商品信息！","product_info_add.asp?bid="&bid&"&sid="&sid&"")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品信息添加</title>
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
		<td colspan="2" class="title">商品信息添加</td>
	</tr>
	<tr>
		<td>商品名称：</td>
		<td><input type="text" name="product_info_name" size="30"></td>
	</tr>
	<tr>
		<td>商品货号：</td>
		<td><input type="text" name="product_info_no" size="30"></td>
	</tr>
	<tr>
		<td>所属类别：</td>
		<td><select name="bid" onChange="changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)">
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
           		<%if sid<>"" then%><option value="<%=prod_info_sid%>" selected><%=prod_SmallClass_name%></option><%end if%> 
         	</select>
		</td>
	</tr>

	<tr>
		<td>市 场 价：</td>
		<td><input type="text" name="product_info_PriceM" size="30"></td>
	</tr>
	<tr>
		<td>本 站 价：</td>
		<td><input type="text" name="product_info_PriceS" size="30"></td>
	</tr>
	<tr>
		<td>库 存 量：</td>
		<td>
		        <input type="text" name="product_info_KuCun" size="30">件</td>
	</tr>
	<tr>
		<td>小 图 片：</td>
		<td>
		        <input type="text" name="product_info_PicS" size="30">
		        <input type="button" value=">>点此上传商品小图片" name="action0" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicS','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')">
		</td>
	</tr>
<tr>
		<td>大 图 片：</td>
		<td>
		        <input type="text" name="product_info_PicB" size="30">
		        <input type="button" value=">>点此上传商品大图片" name="action1" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')">
        		<br>
		<input type="checkbox" name="MorePic" value="1" onClick='showlist(paipai);'>要上传多张商品大图片,请在前面方框内打勾(<font color="#808080">最多共支持三张商品大图片</font>)</td>
	</tr>
	    <tr id=paipai style="display:none">
		<td>多图上传：</td>
		<td>
			第二张商品大图：<input type="text" name="product_info_PicB2" size="30" readonly> <a href="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB2','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=410,height=230')">
			<img src="images/upload.gif" alt="上传图片" style="cursor: hand;" onMouseOver="window.status='使用系统自带的上传程序上传图片';return true;" onMouseOut="window.status='';return true;" border="0"></a><br>
			第三张商品大图：<input type="text" name="product_info_PicB3" size="30" readonly> <a href="javascript:openWin('Njj_Pic_Upload.asp?Fname=product_info_PicB3','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=410,height=230')">
			<img src="images/upload.gif" alt="上传图片" style="cursor: hand;" onMouseOver="window.status='使用系统自带的上传程序上传图片';return true;" onMouseOut="window.status='';return true;" border="0"></a>
		</td>
	</tr>
	<tr>
		<td>详细描述：</td>
		<td>
		<textarea cols=60 rows=20 id="content" name="Content"></textarea>
		</td>
	</tr>
	<tr>
		<td>商品特性：</td>
		<td><input type="checkbox" name="product_info_flag" value="1">新品&nbsp; 
		<input type="checkbox" name="product_info_flag" value="2">推荐&nbsp; 
		<input type="checkbox" name="product_info_flag" value="3">特价</td>
	</tr>
	<tr>
		<td>是否上架：</td>
		<td><input type="radio" value="0" name="product_info_OnOff" checked>上架(显示)&nbsp;&nbsp;
		<input type="radio" value="1" name="product_info_OnOff">下架(隐藏) </td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		<input type="submit" value="提交" name="Submit1">&nbsp;&nbsp;&nbsp; 
		<input type="reset" value="重置" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>

