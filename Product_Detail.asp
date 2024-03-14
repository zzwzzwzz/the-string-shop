<center>
<SCRIPT language="JavaScript" src="js/URLEncode.js"></SCRIPT>
<SCRIPT language="JavaScript" src="js/PicFit.js"></SCRIPT>
<%
dim dbpath,url
dbpath=""
url=request.ServerVariables("Server_NAME")&request.ServerVariables("SCRIPT_NAME") 
if(len(trim(request.ServerVariables("QUERY_STRING")))>0) then 
  url=url & "?" & request.ServerVariables("QUERY_STRING") 
end if 
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file="include/Pages.asp"-->
<!--#include file=Sub.asp -->
<%
dim id
id=my_request("id",1)

//更新商品浏览次数
sql="update Product_info set Product_info_hitnums=Product_info_hitnums+1 where id="&id
conn.execute (sql)

dim Product_info_name,bid,sid,Product_info_PicB,Product_info_flag,Product_info_PriceM,Product_info_PriceS,Product_info_detail,Product_info_kucun
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select product_info_name,bid,sid,product_info_PicB,Product_info_PicB2,Product_info_PicB3,product_info_flag,product_info_PriceM,product_info_PriceS,product_info_detail,Product_info_kucun,Product_info_no from Product_info where id="&id
rs.open sql,conn,1,1
Product_info_name   = rs(0)
bid                 = rs(1)
sid                 = rs(2)
Product_info_PicB   = rs(3)
Product_info_PicB2  = rs(4)
Product_info_PicB3  = rs(5)
Product_info_flag   = rs(6)
Product_info_PriceM = rs(7)
Product_info_PriceS = rs(8)
Product_info_detail = rs(9)
Product_info_kucun  = rs(10)
Product_info_no     = rs(11)
rs.close
set rs=nothing

product_info_prices=FormatNumber(product_info_PriceS,2,-1)
product_info_prices=replace(product_info_prices,",","")

txt=""
if instr(Product_info_flag,1) then txt="推荐、"
if instr(Product_info_flag,2) then txt=txt&"新品、"
if instr(Product_info_flag,3) then txt=txt&"特价"
  
//调出商品大类名称
sql="select prod_BigClass_name from prod_BigClass where prod_BigClass_id="&Bid
set rs=conn.execute (sql)
BClass=rs(0)
rs.close
set rs=nothing

//调出商品小类名称
sql="select prod_SmallClass_name from prod_SmallClass where prod_SmallClass_id="&Sid
set rs=conn.execute (sql)
SClass=rs(0)
rs.close
set rs=nothing

txt_nav="<a href=Product_listCategory.asp?bid="&bid&"> "&Bclass&"</a> &raquo; <a href=Product_listCategory.asp?bid="&bid&"&sid="&sid&">"&SClass&"</a> &raquo; 商品介绍"
action=my_request("action",0)
if action="save" then
    call Product_ReviewAddSave()
end if
%>
<script language="JavaScript" for="window" event="onload">
ImagePreload('<%=Product_info_PicB2%>');
ImagePreload('<%=Product_info_PicB3%>');
</script>
<%
call up(Product_info_name,"商品介绍",txt_nav)
response.write  "<tr>"&_
				"	<td colspan=3><h2 align=center>"&Product_info_name&"</h2></td>"&_
				"</tr>"&_
				"<tr>"&_
				"	<td valign=top align=center width='50%'>"&_
				"    	<a href=uploadpic/"&product_info_PicB&" style='cursor:hand' onclick='OpenFullSizeWindow(ShowImage.src);return false'><img name=ShowImage src=uploadpic/"&product_info_PicB&" onload=fitSize();><br><span id=ShowImgText></span>放大查看</a>"
						if Product_info_PicB2<>"" or Product_info_PicB3<>"" then
response.write "    		<br><a onmouseover=GetShowImg('默认图片','"&product_info_PicB&"');><img src=uploadpic/"&product_info_PicB&" onload='loaded(this,80,80)' style='border: 1px solid #C0C0C0;'></a>&nbsp;"
							if Product_info_PicB2<>"" then
response.write "    			<a onmouseover=GetShowImg('第二张图片','"&product_info_PicB2&"');><img src=uploadpic/"&product_info_PicB2&" onload='loaded(this,80,80)' style='border: 1px solid #C0C0C0;'></a>&nbsp;"
							end if
							if Product_info_PicB3<>"" then
response.write "    			<a onmouseover=GetShowImg('第三张图片','"&product_info_PicB3&"');><img src=uploadpic/"&product_info_PicB3&" onload='loaded(this,80,80)' style='border: 1px solid #C0C0C0;'></a>&nbsp;"
							end if
						end if
response.write  "   </td>"&_
				"	<td valign=top width='35%'>"&_
				"		<table border=0 width=100% cellpadding=3 style=border-collapse: collapse>"&_
				"			<tr><td>所属类别： <a href=Product_listCategory.asp?bid="&bid&"> "&Bclass&"</a> &raquo; <a href=Product_listCategory.asp?bid="&bid&"&sid="&sid&">"&SClass&"</a></td></tr>"
							if product_info_no<>"" then
response.write  "				<tr><td>商品货号： "&product_info_no&"</td></tr>"
							end if
response.write  "			<tr><td>商品特性： "&txt&"</td></tr>"&_
				"			<tr><td>市 场 价： <font color=#808080>￥"&FormatNumber(product_info_PriceM,2,-1)&"</font></td></tr>"&_
				"			<tr><td>本 站 价： <font color=#FF6600 size=4> <b>￥"&product_info_prices&"</b></font></td></tr>"
							if product_info_kucun<>"" then
response.write  "			<tr><td>库存情况： "
								if product_info_kucun>0 then 
									response.write "有货"
								else 
									response.write "缺货中"
								end if
response.write  "			</td></tr>"
							end if
response.write  "			<tr><td style='border-bottom: 1px solid #E8E8E8'>"
							if product_info_kucun<>"" and product_info_kucun<=0 then
								response.write "缺货中,暂时不能下单"
							else 
								response.write  "<a href=Cart_Add.asp?id="&id&"><img src=images/add_shop_cart.gif></a>&nbsp;&nbsp;<a href=Product_Favorite.asp?id="&id&"><img src=images/add_shop_fav.gif ></a>"
							end if
response.write  "			</td></tr>"
response.write  "		</table>"&_
				"</td>"&_
				"<td valign=top width='20%'>"
response.write  "	</td>"&_
				"</tr>"&_
 				"<tr><td colspan=3 class=RightHead>商品详细描述</td></tr>"&_
				"<tr><td colspan=3 style=table-layout:fixed;word-break:break-all class=maintxt>"&product_info_detail&"<br><br></td></tr>"&_
				"</table>"
				set rs=server.createobject("adodb.recordset")
        		sql="select prod_review_detail,prod_review_name,prod_review_time,prod_review_backdetail,prod_review_BackTime from prod_review where prod_review_pid="&id&" order by prod_review_time desc"
        		rs.open sql,conn,1,1
        		iCount=rs.RecordCount '记录总数 
response.write  "<table width='100%' cellspacing=0 cellpadding=3 style='border-collapse: collapse' class=righttable><tbody class=table_td>"&_
				"<tr><td colspan=2 class=RightHead>商品评论及咨询</td></tr>"
       			if rs.eof then 
            		response.write "<tr><td colspan=2 align=center>目前暂无此商品相关评价及咨询信息,<a href=#review onClick=showlist('a');>点此发表您的评论或咨询信息</a>！</td></tr>"
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

        			    set prod_review_detail    =rs(0)
            			set prod_review_name      =rs(1)
            			set prod_review_time      =rs(2)
            			set prod_review_BackDetail=rs(3)
            			set prod_review_BackTime  =rs(4)
        
            			while not rs.eof and i<=rs.pagesize 
        			
        			    response.write  "<tr><td valign=top width=8%  style='padding-top:8px;'>&nbsp;作者：</td><td style='padding-top:8px;'> "&prod_review_name&"&nbsp;&nbsp;发表时间： "&prod_review_time&"</td></tr>"&_
	    			    				"<tr><td valign=top width=8% >&nbsp;内容：</td><td>"&prod_review_detail&"</td></tr>"
	    			    if prod_review_backdetail<>"" then 
	    			        response.write "<tr><td valign=top width=8% >&nbsp;<font color='#FF6600'><b>回复：</font></b></td><td><font color='#FF6600'>"&prod_review_BackDetail&"</font></td></tr>"
	    			    end if
	    			    response.write "<tr><td colspan=2><hr></td></tr>"
	    			    
	    			    rs.movenext
        				i=i+1
            			wend
            			response.write "<tr><td colspan=2>"
            			call PageControl(iCount,maxpage,page)
            			response.write "</td></tr>"
            			response.write "<tr><td colspan=2 align=center><a href=#review onClick=showlist('a');>点此发表您的评论或咨询信息</a></td></tr>"
        			end if
        			rs.close
        			set rs=nothing
response.write  "</tbody></table>"&_
				"<table width='100%' cellspacing=0 cellpadding=4 style='border-collapse: collapse' border=0 id=linkimg style='display:none' class=righttable><tbody class=table_td>"&_
				"<form action=Product_Detail.asp method=post name=form1>"&_
				"<input type=hidden name=action value=save>"&_
				"<input type=hidden name=prod_review_pid value="&id&">"&_
				"<tr><td colspan=2 class=RightHead>发表您的评论或咨询信息</td></tr>"&_
				"<tr><td>您的称呼：</td><td>"
				if session("user_info_LoginIn")=true then
response.write  "	<input type=text name=prod_review_name size=25 value="&session("user_info_username")&">"
				else
response.write  "	<input type=text name=prod_review_name size=25>"
				end if
response.write  "</td></tr>"&_
				"<tr><td valign=top>评论或咨询：</td><td><textarea rows=6 name=prod_review_detail cols=50></textarea></td></tr>"&_
				"<tr><td></td><td><input class=button type=submit value=提交信息><br><br></td></tr>"&_
				"</form>"
        			
call down()
%>
</center>