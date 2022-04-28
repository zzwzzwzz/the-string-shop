<%
dim url
url=request.ServerVariables("SCRIPT_NAME") 
if(len(trim(request.ServerVariables("QUERY_STRING")))>0) then 
  url=url & "?" & request.ServerVariables("QUERY_STRING") 
end if

user_info_id=session("user_info_id")
if session("user_info_id")<>"" then
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select user_info_mark from user_info where user_info_id="&user_info_id
	rs.open sql,conn,1,1
	user_info_mark=rs(0)
	rs.close
	set rs=nothing

    sql="select user_level_Name,user_level_rebate from user_Level where user_level_markmin<="&user_info_mark&" and user_level_markmax>="&user_info_mark&""
  	set rs=conn.execute (sql)
  	user_level_Name=rs(0)
  	user_level_rebate=rs(1)
  	rs.close
  	set rs=nothing
end if

'调出会员登陆框显示否的选项
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_OnOffIndexUserLogin from root_option where id=1"
rs.open sql,conn,1,1
root_option_OnOffIndexUserLogin=rs(0)
rs.close
set rs=nothing
if root_option_OnOffIndexUserLogin=0 then

	//<!----member login or reg  ---->		
	response.write "<table width='100%' cellspacing=1 cellpadding=2 class=MainTable><tbody class=table_td>"
				if session("user_info_id")<>"" and session("user_info_LoginIn")=true then
	response.write  "	<tr><td class=MainHead>我的帐户</td></tr>"&_
				"	<tr><td>欢迎您:<b><font color=#FFB5B5>"&session("user_info_UserName")&"</font></b></td></tr>"&_
				"	<tr><td><a href=User_Personal.asp>基本资料</a></td></tr>"&_
				"	<tr><td><a href=User_OrderList.asp>我的订单</a></td></tr>"&_
				"	<tr><td><a href=User_fav.asp>商品收藏</a></td></tr>"&_
				"	<tr><td><a href=User_LoginOut.asp>[退出登录]</a></td></tr>"
				else
	response.write  "	<form name=form_login action=User_loginCheck.asp method=post onsubmit='return submit1();'>"&_
				"	<input type=hidden name=urlpath value="&url&">"&_
				"	<tr><td colspan=2 class=MainHead>会员登陆/注册</td></tr>"&_
				"	<tr><td>&nbsp;用户名：<input type=text size=14 name=loginname></td></tr>"&_
				"	<tr><td>&nbsp;密　码：<input type=password size=14 name=loginpass></td></tr>"&_
				"	<tr><td>&nbsp;验证码：<input type=text size=7 name=codeid>&nbsp;<img src=Include/checkcode.asp></td></tr>"&_
				"	<tr><td align=center>&nbsp;<input class=button type=submit value=' 登 陆 '>  <a href=User_PassWordGet.asp>忘记密码</a></td></tr>"&_
				"	<tr><td align=center><input class=button type=button value=立即注册 onclick=window.location='User_Reg.asp'></td></tr>"&_
				"	</form>"
				end if
	response.write  "<table width='100%' cellspacing=1 cellpadding=2 class=MainTable></table>"&_
				"<div class=brclass></div>"
end if
//<!----product class  ---->
set rs=server.createobject("adodb.recordset")
sql="select root_option_NumsPerRowSclass from root_option where id=1"
rs.open sql,conn,1,1
root_option_NumsPerRowSclass=rs(0)
rs.close
set rs=nothing
if root_option_NumsPerRowSclass=2 then		
	response.write  "<table width='100%' cellspacing=1 cellpadding=2 class=MainTable></table>"&_
					"<div class=brclass></div>"
	response.write  "<table width='100%' cellspacing=0 cellpadding=4 class=category_table>"&_
					"<tr><td class=MainHead colspan=2>商品分类</td></tr>"
					Set rs= Server.CreateObject("ADODB.Recordset")
					sql="select prod_BigClass_id,prod_BigClass_name from prod_BigClass order by prod_BigClass_sort asc"
					rs.open sql,conn,1,1
					if not rs.eof then
    					set prod_BigClass_id=rs(0)
    					set prod_BigClass_name=rs(1)
    					while not rs.eof
    						response.write "<tr><td colspan=2><img src=images/icon_arrow_blue.gif> <a href=Product_ListCategory.asp?Bid="&prod_BigClass_id&" class=left_bid><b>"&prod_BigClass_Name&"</b></a></td></tr>"
    						//调出小类别
    						set rs_s=server.CreateObject("adodb.recordset")
							sql_s="select prod_SmallClass_id,prod_SmallClass_name,prod_SmallClass_bid from prod_SmallClass where prod_SmallClass_Bid=" & prod_BigClass_id & " order by prod_SmallClass_id"
    						rs_s.open sql_s,conn,1,1
    						if not rs_s.eof then
        						set prod_SmallClass_id=rs_s(0)
        						set prod_SmallClass_name=rs_s(1)
        						set prod_SmallClass_bid=rs_s(2)
        						i=1
       							while not rs_s.eof
        						response.write "<td>&nbsp;&nbsp;<a href=Product_ListCategory.asp?Bid="&prod_SmallClass_Bid&"&Sid="&prod_SmallClass_id&">"&prod_SmallClass_name&"</a></td>"
	    						if (i mod 2)=0 then
	    							response.write "</tr>"
	  							end if
	  							rs_s.movenext
	  							i=i+1
	    						wend
							end if
							rs_s.close
							set rs_s=nothing
						rs.movenext
						wend
					end if
					rs.close
					set rs=nothing 
	response.write  "</table>"&_
			"<div class=brclass></div>"
else
	response.write  "<table width='100%' cellspacing=1 cellpadding=4 class=MainTable><tbody class=table_td>"&_
					"	<tr><td class=MainHead>商品分类</td></tr>"
					Set rs= Server.CreateObject("ADODB.Recordset")
					sql="select prod_BigClass_id,prod_BigClass_name from prod_BigClass order by prod_BigClass_sort asc"
					rs.open sql,conn,1,1
					if not rs.eof then
    					set prod_BigClass_id=rs(0)
    					set prod_BigClass_name=rs(1)
    					while not rs.eof
    						response.write "<tr><td><img src=images/icon_arrow_blue.gif> <a href=Product_ListCategory.asp?Bid="&prod_BigClass_id&" class=left_bid><b>"&prod_BigClass_Name&"</b></a></td></tr>"
    						//调出小类别
    						set rs_s=server.CreateObject("adodb.recordset")
							sql_s="select prod_SmallClass_id,prod_SmallClass_name,prod_SmallClass_bid from prod_SmallClass where prod_SmallClass_Bid=" & prod_BigClass_id & " order by prod_SmallClass_id"
    						rs_s.open sql_s,conn,1,1
    						if not rs_s.eof then
        						set prod_SmallClass_id=rs_s(0)
        						set prod_SmallClass_name=rs_s(1)
        						set prod_SmallClass_bid=rs_s(2)
       							while not rs_s.eof
        						response.write "<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Product_ListCategory.asp?Bid="&prod_SmallClass_Bid&"&Sid="&prod_SmallClass_id&">"&prod_SmallClass_name&"</a></td></tr>"
	  							rs_s.movenext
	    						wend
							end if
							rs_s.close
							set rs_s=nothing
						rs.movenext
						wend
					end if
					rs.close
					set rs=nothing 
	response.write  "</tbody></table>"&_
					"<div class=brclass></div>"
end if

//<!----hot top10  ---->
response.write  "<table width=100% cellspacing=0 cellpadding=4 class=MainTable>"&_
				"	<tbody class=table_td><tr>"&_
				"		<td class=MainHead>最新文章</a></td>"&_
				"		<td class=MainHead align=right><a href=News_List.asp class=U><span style='font-weight: 200'>更多文章</span></a></td>"&_
				"	</tr>"
'最新10条文章调出
set rs = server.createobject("ADODB.Recordset")
sql = "select top 10 id, news_info_title, news_info_type, news_info_content from news_info order by id desc"
rs.open sql, conn, 1, 1
if not rs.eof then
    dim news_info_title, news_info_content, news_info_addtime, news_info_hitnums
    set news_info_id      = rs(0)
    set news_info_title   = rs(1)
    set news_info_type    = rs(2)
    set news_info_content = rs(3)
    while not rs.eof
        if len(news_info_title) > 22 then 
            news_info_title1 = left(news_info_title, 20)&"..."
        else
            news_info_title1 = news_info_title
        end if
        response.write "<tr><td colspan=2>"

        if news_info_type = 1 then
            response.write "·<a href=News_Detail.asp?id="&news_info_id&">"&news_info_title1&"</a>"
        else
            response.write "·<a href="&news_info_content&" target=_blank>"&news_info_title1&"</a>"
        end if
        response.write "</td></tr>"
        rs.movenext
    wend
end if
rs.close
set rs=nothing
response.write "</tbody></table>"&_
				"<div class=brclass></div>"


//<!----  vote  ---->
set rs=server.createobject("adodb.recordset")
sql="select base_vote_OnOff from base_vote where base_vote_flag=1"
rs.open sql,conn,1,1
base_vote_OnOff=rs(0)
rs.close
set rs=nothing

if base_vote_OnOff=0 then
	response.write  "<table width=100% cellspacing=1 cellpadding=4 class=MainTable><tbody class=table_td>"&_
					"<tr><td class=MainHead>投票调查</td></tr>"&_
					"<form action=votes.asp?vflag=add method=post target=win onSubmit=windowOpener()>"
   				 	sql="select base_vote_detail from base_vote where base_vote_flag=1"
    			 	set rs=conn.execute (sql)
    			 	base_vote_title=rs(0)
    			 	rs.close
    			 	set rs=nothing
	response.write  "<tr><td align=left>"&base_vote_title&"</td></tr>"&_
    				"<tr><td>"
    				sql="select base_vote_id,base_vote_detail from base_vote where base_vote_flag=0"
    				set rs=conn.execute (sql)
    				if not rs.eof then
        				set base_vote_id=rs(0)
        				set base_vote_detail=rs(1)
        				do while not rs.eof
	response.write  "	<input type=radio value="&base_vote_id&" name=idnums>"&base_vote_detail&"<br>"
        				rs.movenext
        				loop
    				end if
    				rs.close
    				set rs=nothing
	response.write  "</td></tr>"&_
    				"<tr><td align=center><input class=button type=submit value=投票及查看结果></td></tr>"&_
    				"</form></tbody></table><div class=brclass></div>"
end if
%>
