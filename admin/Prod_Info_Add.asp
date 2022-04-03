<!--#include file="admin_check.asp"-->
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品信息添加</title>
<link rel="stylesheet" type="text/css" href="style.css">
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
<script src="Editor/edit.js" type="text/javascript"></script>
</head>

<body>
<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="prod_Info_AddSave.asp" method="post" name="form1" encType="multipart/form-data">
<input type="hidden" name="filepath" value="../uploadpic/"> 
    <tr>
		<td colspan="2" class="title">商品信息添加</td>
	</tr>
	<tr>
		<td>商品名称：</td>
		<td><input type="text" name="prod_info_name" size="30"></td>
	</tr>
	<tr>
		<td>商品货号：</td>
		<td><input type="text" name="prod_info_no" size="15"></td>
	</tr>
	<tr>
		<td>加促销语：</td>
		<td><select size="1" name="prod_info_AdWord">
		<option value="" selected>不加任何促销语(默认)</option>
		<option value="赞">赞</option>
		<option value="热门">热门</option>
		<option value="店长赞">店长赞</option>
		<option value="抢购中">抢购中</option>
		<option value="价格新底">价格新底</option>
		<option value="店长严重推荐">店长严重推荐</option>
		<option value="新品上市，棒！">新品上市，棒！</option>
		<option value="超特价，当到谷底">超特价，当到谷底</option>
		<option value="人气绝项,销售佳！">人气绝项,销售佳！</option>
		</select></td>
	</tr>
	<tr>
		<td>所属类别：</td>
		<td><select name="cid" size="5">
		    <%
		     sql="select cid,prod_class_name from prod_class order by cid desc"
		     set rs=conn.execute (sql)
		     set cid=rs(0)
		     set prod_class_name=rs(1)
		     do while not rs.eof
		    %>
		    <option value="<%=cid%>" <%if cid=cint(request("cid")) then response.write "selected" %>><%=prod_class_name%></option>
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
		<td><input type="text" name="prod_info_PriceM" size="10"> 元</td>
	</tr>
	<tr>
		<td>本 站 价：</td>
		<td><input type="text" name="prod_info_PriceS" size="10"> 元</td>
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
								<td id="previewImage2" valign="middle" align="center" style="width:100px;height:100px;border:solid 1px #DDD">没有图片</td>
							</tr>
						</table>
						
					</td>
				</tr>
			</table>
		</td>
	</tr>	
	<%end if%>

	<tr>
		<td>商品大图：</td>
		<td>
			<table border="0" width="100%" cellpadding="2" style="border-collapse: collapse">
				<tr>
					<td>
						<input type="file"  size="30" name="file"  onchange="document.getElementById('previewImage').innerHTML = '<img src=\''+this.value+'\' width=100 height=100  onload=\'ResizeImage(this, 100, 100);\' align=absmiddle>';" /> 
						<p><font color="#666666"><span class="posttip">
						图片限定120k内，jpg或gif格式，请确保图片在浏览器中可以正常打开。<br>
						一张好图胜千言，建议500×500象素，主体突出。<br>
						</span></font><font color="#FF3300">
						上传大图，即可自动生成商品缩略图！</font>
					</td>
					<td>
						<table border="0" cellspacing="0" cellpadding="0" width="100" height="100">
							<tr>
								<td id="previewImage" valign="middle" align="center" style="width:100px;height:100px;border:solid 1px #DDD">没有图片</td>
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
           document.write ('<iframe src="prod_Txtbox.asp" id="message" width="90%" height="300"></iframe>')
           frames.message.document.designMode = "On";
           </script>
		</td>
	</tr>
	<tr>
		<td>商品特性：</td>
		<td><input type="checkbox" name="prod_info_flag" value="1">新品&nbsp; 
		<input type="checkbox" name="prod_info_flag" value="2">推荐&nbsp; 
		<input type="checkbox" name="prod_info_flag" value="3">特价</td>
	</tr>
	<tr>
		<td>是否上架：</td>
		<td><input type="radio" value="0" name="prod_info_OnOff" checked>上架(显示)&nbsp;&nbsp;
		<input type="radio" value="1" name="prod_info_OnOff">下架(隐藏) </td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		<input type="submit" value="提交(按enter键也可以快速提交)" name="Submit1" onclick="document.form1.Content.value = frames.message.document.body.innerHTML;">&nbsp;&nbsp;&nbsp; 
		    <input type="reset" value="重置" name="B2">
		    <input type="hidden" name="Content" value>
		</td>
	</tr>
</form>
</tbody>
</table>
</body>

</html>

