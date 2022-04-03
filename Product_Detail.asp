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
sql="select product_info_name,bid,sid,product_info_PicB,Product_info_PicB2,Product_info_PicB3,product_info_flag,product_info_PriceM,product_info_PriceS,product_info_detail,Product_info_kucun,Product_info_no,Product_info_brand from Product_info where id="&id
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
Product_info_brand  = rs(12)
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

//调出商品品牌名称
sql="select prod_brand from prod_brand where id="&Product_info_brand
set rs=conn.execute (sql)
prod_brand=rs(0)
rs.close
set rs=nothing

txt_nav="<a href=Product_listCategory.asp?bid="&bid&"> "&Bclass&"</a> &raquo; <a href=Product_listCategory.asp?bid="&bid&"&sid="&sid&">"&SClass&"</a> &raquo; 商品介绍"

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_info_QQ,root_info_QQOnOff,root_info_WangWang,root_info_WangWangOnOff from root_info where id=1"
rs.open sql,conn,1,1
root_info_QQ            =rs(0)
root_info_QQOnOff       =rs(1)
root_info_WangWang      =rs(2)
root_info_WangWangOnOff =rs(3)
rs.close
set rs=nothing

'调出不同会员级别价格的显示方式 / 积分与价格换算值
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_MarkYuan,root_option_PriceShowType from root_option where id=1"
rs.open sql,conn,1,1
root_option_MarkYuan       = rs(0)
root_option_PriceShowType  = rs(1)
rs.close
set rs=nothing
x=1/root_option_MarkYuan
 	
select case root_option_PriceShowType
	case 0 '不显示
	 	pricetxt=""
	case 2 '全显示
		'调出会员级别与折扣列表
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select user_level_Name,user_level_rebate from user_Level order by user_level_markmin asc"
		set rs=conn.execute (sql)
		rs.open sql,conn,1,1
		set user_level_Name=rs(0)
		set user_level_rebate=rs(1)
		while not rs.eof
		xxx=user_level_rebate/100
		yyy=Product_info_PriceS*xxx
		pricetxt1=pricetxt1&user_level_Name&":"&FormatNumber(yyy,2,-1)&"&nbsp;"
		rs.movenext
		wend
		rs.close
		set rs=nothing
		pricetxt="<tr><td>"&pricetxt1&"</td></tr>"
	case 1 '会员登陆后显示同级及以下级会员价
		if session("user_info_id")<>"" then
			user_info_id=session("user_info_id")
			Set rs= Server.CreateObject("ADODB.Recordset")
			sql="select user_info_mark from user_info where user_info_id="&user_info_id
			rs.open sql,conn,1,1
			user_info_mark=rs(0)
			rs.close
			set rs=nothing
	
			sql="select user_level_name,user_level_rebate from user_Level where user_level_markmin<="&user_info_mark&" and user_level_markmax>="&user_info_mark&""
			set rs=conn.execute (sql)
			user_level_name=rs(0)
			user_level_rebate=rs(1)
			rs.close
			set rs=nothing
  	
			m=Product_info_PriceS*user_level_rebate/100
			pricetxt="<tr><td>"&user_level_Name&"价:"&FormatNumber(m,2,-1)&"</td></tr>"
		else
			pricetxt="<tr><td>"
			'调出会员级别与折扣列表
			Set rs=Server.CreateObject("ADODB.Recordset")
			sql="select user_level_Name from user_Level order by user_level_markmin asc"
			set rs=conn.execute (sql)
			rs.open sql,conn,1,1
			set user_level_Name=rs(0)
			while not rs.eof
			pricetxt1=pricetxt1&"<li>"&user_level_Name&"价:会员登陆后查看</li>"
			rs.movenext
			wend
			rs.close
			set rs=nothing
			pricetxt="<tr><td>"&pricetxt1&"</td></tr>"
        end if
end select
			

'不同级别会员享受积分计算
user_info_id=session("user_info_id")
if session("user_info_id")<>"" then
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select user_info_mark from user_info where user_info_id="&user_info_id
	rs.open sql,conn,1,1
	user_info_mark=rs(0)
	rs.close
	set rs=nothing
	
	sql="select user_level_rebate from user_Level where user_level_markmin<="&user_info_mark&" and user_level_markmax>="&user_info_mark&""
	set rs=conn.execute (sql)
	user_level_rebate=rs(0)
	rs.close
	set rs=nothing
  	
  	'打折时积分
	m=Product_info_PriceS*user_level_rebate/100
	y=m/x
	y=cint(y)
else
	'不打折时积分
	y=Product_info_PriceS/x
	y=cint(y)
end if

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
							if product_info_brand<>"" then
response.write  "				<tr><td>商品品牌： "&prod_brand&"</td></tr>"
							end if
response.write  "			<tr><td>商品特性： "&txt&"</td></tr>"&_
				"			<tr><td>市 场 价： <font color=#808080>￥"&FormatNumber(product_info_PriceM,2,-1)&"</font></td></tr>"&_
				"			<tr><td>本 站 价： <font color=#FF6600 size=5> <b>￥"&product_info_prices&"</b></font></td></tr>"&_
				pricetxt&_
				"			<tr><td>赠送积分： <font color=#FF6600> "&y&"</font></td></tr>"
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
		              		Set rs=Server.CreateObject("ADODB.Recordset")
		              		sql="select root_option_OnOffAliPayButton from root_option where id=1"
		              		rs.open sql,conn,1,1
                      		root_option_OnOffAliPayButton=rs(0)
                      		rs.close
                      		set rs=nothing
                      		if root_option_OnOffAliPayButton=1 then
                      			if product_info_kucun<>"" and product_info_kucun>0 then
response.write "    				<tr><td><a href=OnlyOne_ByAlipay.asp?flag=1&url="&url&"&product_info_PriceS="&product_info_PriceS&"&product_info_name="&trim(product_info_name)&" target=_blank><img src=images/zhifubao.gif width=100></a></td></tr>"
							    end if
							end if
response.write  "		</table>"&_
				"</td>"&_
				"<td valign=top width='15%'>"
						if (root_info_QQOnOff=0 and root_info_QQ<>"") or (root_info_WangWangOnOff=0 and root_info_WangWang<>"") then
response.write "		<table width=100% cellspacing=0 cellpadding=4 class=MainTable><tbody class=table_td>"
							if root_info_QQOnOff=0 and root_info_QQ<>"" then
response.write "    		<tr><td class=mainhead>QQ咨询：</td></tr>"
								qq=split(root_info_QQ,",")
                            	for k=0 to ubound(qq)  
response.write "    				<tr><td><a target=_blank href=http://wpa.qq.com/msgrd?V=1&Uin="&trim(qq(k))&"&Site=购物咨询&Menu=yes><img src=http://wpa.qq.com/pa?p=1:"&trim(qq(k))&":16 alt=QQ咨询></a></td></tr>"
								next
							end if
							if root_info_WangWangOnOff=0 and root_info_WangWang<>"" then
response.write  "			<tr>"&_
				"				<td>淘宝旺旺："
%>								<script language="javascript">
						        var taobaoid;
						        var taobaos;
                                taobaos="<%=root_info_WangWang%>";
                                taobaoid=URLEncode(taobaos)
                                document.writeln("<a target=_blank href=http://amos1.taobao.com/msg.ww?v=2&s=1&uid="+taobaoid+">")
						        document.writeln("<img border=0 alt=点击这里给我发消息 src=http://amos1.taobao.com/online.ww?v=2&s=1&uid="+taobaoid+">")
						        document.writeln("</a>")
						    	</script>
<%
response.write  "				</td>"&_
				"			</tr>"
							end if
response.write  "		</tbody></table>"
						end if
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
            		response.write "<tr><td colspan=2 align=center>目前暂无此商品相关评价及咨询信息,<a href=#review onClick=showlist('a');><img src=images/ico_add_guestbook.gif>点此发表您的评论或咨询信息</a>！</td></tr>"
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
            			response.write "<tr><td colspan=2 align=center><a href=#review onClick=showlist('a');><img src=images/ico_add_guestbook.gif>点此发表您的评论或咨询信息</a></td></tr>"
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
				"<tr><td>验证码：</td><td><input name=codeid size=10><img src=Include/checkcode.asp> 请照样输入彩色数字(验证码)</td></tr>"&_
				"<tr><td></td><td><input class=button type=submit value=提交信息><br><br></td></tr>"&_
				"</form>"
        			
call down()

dim endtime
endtime=timer()
response.write "页面执行时间："&FormatNumber((endtime-startime)*1000,3)&"毫秒"
%>
</center>