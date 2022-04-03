<!--#include file="admin_check.asp"-->
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""参数错误!"");location.href=""prod_info_List.asp"";</script>")
  response.end
end if

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select id,cid,prod_info_name,prod_info_flag,prod_info_PriceM,prod_info_PriceS,prod_info_PicB,prod_info_PicS,prod_info_OnOff,prod_info_AdWord,prod_info_no from prod_info where id="&id
rs.open sql,conn,1,1
id				=rs(0)
cid				=rs(1)
prod_info_name	=rs(2)
prod_info_flag	=rs(3)
prod_info_PriceM=rs(4)
prod_info_PriceS=rs(5)
prod_info_PicB	=rs(6)
prod_info_PicS	=rs(7)
prod_info_OnOff =rs(8)
prod_info_AdWord=rs(9)
prod_info_no    =rs(10)
rs.close
set rs=nothing
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品信息编辑</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
<script language = "JavaScript">
var imageObject;
function ResizeImage(obj, MaxW, MaxH)
{
    if (obj != null) imageObject = obj;
    var state=imageObject.readyState;
    var oldImage = new Image();
    oldImage.src = imageObject.src;
    var dW=oldImage.width; var dH=oldImage.height;
    if(dW>MaxW || dH>MaxH) {
        a=dW/MaxW; b=dH/MaxH;
        if(b>a) a=b;
        dW=dW/a; dH=dH/a;
    }
    if(dW > 0 && dH > 0)
        imageObject.width=dW;imageObject.height=dH;
    if(state!='complete' || imageObject.width>MaxW || imageObject.height>MaxH) {
        setTimeout("ResizeImage(null,"+MaxW+","+MaxH+")",40);
    }
}

function findItem(n, d) {
	var p,x,i;
	if(!d) d=document;
	if((p=n.indexOf("?"))>0&&parent.frames.length) {
		d=parent.frames[n.substring(p+1)].document;
		n=n.substring(0,p);
	}
	if(!(x=d[n])&&d.all)
		x=d.all[n];
	for (i=0;!x&&i<d.forms.length;i++)
		x=d.forms[i][n];
	for(i=0;!x&&d.layers&&i<d.layers.length;i++)
		x=findItem(n,d.layers[i].document);
	return x;
}

</script>
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="Prod_Info_ModiSave.asp" method="post" name="form1" encType="multipart/form-data">
<input type="hidden" name="filepath" value="../uploadpic/"> 
<input type="hidden" name="id" value="<%=id%>"> 
    <tr>
		<td colspan="2" class="title">商品信息编辑</td>
	</tr>
	<tr>
		<td>商品名称：</td>
		<td>
		<input type="text" name="prod_info_name" size="40" value="<%=prod_info_name%>"></td>
	</tr>
	<tr>
		<td>商品货号：</td>
		<td>
		<input type="text" name="prod_info_no" size="20" value="<%=prod_info_no%>"></td>
	</tr>
	<tr>
		<td>加促销语：</td>
		<td><select size="1" name="prod_info_AdWord">
		<option value="" <%if prod_info_flag="" then response.write "selected"%>>不加任何促销语(默认)</option>
		<option value="赞" <%if prod_info_flag="赞" then response.write "selected"%>>赞</option>
		<option value="热门" <%if prod_info_flag="热门" then response.write "selected"%>>热门</option>
		<option value="店长赞" <%if prod_info_flag="店长赞" then response.write "selected"%>>店长赞</option>
		<option value="抢购中" <%if prod_info_flag="抢购中" then response.write "selected"%>>抢购中</option>
		<option value="价格新底" <%if prod_info_flag="价格新底" then response.write "selected"%>>价格新底</option>
		<option value="店长严重推荐" <%if prod_info_flag="店长严重推荐" then response.write "selected"%>>店长严重推荐</option>
		<option value="新品上市，棒！" <%if prod_info_flag="新品上市，棒！" then response.write "selected"%>>新品上市，棒！</option>
		<option value="超特价，当到谷底" <%if prod_info_flag="超特价，当到谷底" then response.write "selected"%>>超特价，当到谷底</option>
		<option value="人气绝项,销售佳！" <%if prod_info_flag="人气绝项,销售佳！" then response.write "selected"%>>人气绝项,销售佳！</option>
		</select></td>
	</tr>
	<tr>
		<td>所属类别：</td>
		<td><select name="cid">
		    <%
		     sql="select cid,prod_class_name from prod_class order by cid desc"
		     set rs=conn.execute (sql)
		     set cid1=rs(0)
		     set prod_class_name=rs(1)
		     do while not rs.eof
		    %>
		    <option value="<%=cid1%>" <%if cid1=cint(cid) then response.write "selected" %>><%=prod_class_name%></option>
		    <%
		     rs.movenext
		     loop
		     rs.close
		     set rs=nothing
		    %>
		 </select></td>
	</tr>
	<tr>
		<td>市 场 价：</td>
		<td>
		<input type="text" name="prod_info_PriceM" size="10" value="<%=prod_info_PriceM%>"></td>
	</tr>
	<tr>
		<td>本 站 价：</td>
		<td>
		<input type="text" name="prod_info_PriceS" size="10" value="<%=prod_info_PriceS%>"></td>
	</tr>
	<%
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select top 1 root_option_PicSType from root_option"
	rs.open sql,conn,1,1
	root_option_PicSType=rs(0)
	rs.close
	set rs=nothing
	if root_option_PicSType=1 then
	%>
	<tr>
		<td>上传小图：</td>
		<td>
			<table border="0" width="100%" cellpadding="2" style="border-collapse: collapse">
				<tr>
					<td>
						<input type="file"  size="30" name="file2"  onchange="document.getElementById('previewImage2').innerHTML = '<img src=\''+this.value+'\' width=100 height=100  onload=\'ResizeImage(this, 100, 100);\' align=absmiddle>';" /> 
						<p><font color="#666666"><span class="posttip">
						图片限定120k内，jpg或gif格式，请确保图片在浏览器中可以正常打开。<br>
						一张好图胜千言，建议500×500象素，主体突出。</span></font></td>
					<td>
						<table border="0" cellspacing="0" cellpadding="0" width="100" height="100">
							<tr>
								<td id="previewImage2" valign="top" align="center" style="width:100px;height:100px;border:solid 1px #DDD">
								<img src=../uploadpic/<%=prod_info_pics%> width=100 height=100  onload='ResizeImage(this, 100, 100);' align=middle></td>
							</tr>
						</table>
						
					</td>
				</tr>
			</table>
		</td>
	</tr>	
	<%end if%>

	<tr>
		<td>上传大图：</td>
		<td>
			<table border="0" width="100%" cellpadding="2" style="border-collapse: collapse">
				<tr>
					<td>
						<input type="file"  size="30" name="file"  onchange="document.getElementById('previewImage').innerHTML = '<img src=\''+this.value+'\' width=100 height=100  onload=\'ResizeImage(this, 100, 100);\' align=absmiddle>';" /> 
						<p><font color="#666666"><span class="posttip">
						图片限定120k内，jpg或gif格式，请确保图片在浏览器中可以正常打开。<br>
						一张好图胜千言，建议500×500象素，主体突出。<br>
						</span></font><font color="#FF3300">
						上传大图，即可自动生成商品缩略图</font></td>
					<td>
						<table border="0" cellspacing="0" cellpadding="0" width="100" height="100">
							<tr>
								<td id="previewImage" valign="middle" align="center" style="width:100px;height:100px;border:solid 1px #DDD">
								<img src=../uploadpic/<%=prod_info_picB%> width=100 height=100  onload='ResizeImage(this, 100, 100);' align=middle></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>详细描述：</td>
		<td>　
		   <!--//商品介绍//-->
		   <!--#include file="editor/editor.asp"-->
           <script language="javascript">
           document.write ('<iframe src="prod_txtbox.asp?id=<%=id%>&action=modify" id="message" width="90%" height="300"></iframe>')
           frames.message.document.designMode = "On";
           </script>
		</td>
	</tr>
	<tr>
		<td>商品特性：</td>
		<td><input type="checkbox" name="prod_info_flag" value="1" <%if instr(prod_info_flag,1) then response.write "checked" %>>新品&nbsp; 
		<input type="checkbox" name="prod_info_flag" value="2" <%if instr(prod_info_flag,2) then response.write "checked" %>>推荐&nbsp; 
		<input type="checkbox" name="prod_info_flag" value="3" <%if instr(prod_info_flag,3) then response.write "checked" %>>特价</td>
	</tr>
	<tr>
		<td>是否上架：</td>
		<td><input type="radio" value="0" name="prod_info_OnOff" <%if prod_info_OnOff=0 then response.write "checked" %>>上架(显示)&nbsp;&nbsp;
		    <input type="radio" value="1" name="prod_info_OnOff" <%if prod_info_OnOff=1 then response.write "checked" %>>下架(隐藏) </td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="提交" name="Submit1" onclick="document.form1.Content.value = frames.message.document.body.innerHTML;">&nbsp; 
		    <input type="reset" value="重置" name="B2">
		    <input type="hidden" name="Content" value>
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>
