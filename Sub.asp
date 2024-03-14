<%
'***********************************************
' 目的： 公共上部分
' 输入： title,nowplace,nav
' 说明： title:当前页标题; nowplace:当前位置名称; nav:导航
'***********************************************
sub up(title,nowplace,nav)
	response.write "<html>"&_
	"<head>"&_
	"<meta http-equiv=Content-Language content=zh-cn>"&_
	"<meta http-equiv=Content-Type content=text/html; charset=gb2312>"&_
	"<title>"&title&"</title>"&_
	"</head>"&_
	"<body>"
	%>
	<!--#include file=Top.asp-->
	<%
	response.write "<table border=0 width=100% cellpadding=0 style=border-collapse: collapse>"&_
	          	   "	<tr>"&_
				   "		<td width=190 valign=top>"
	%>
	<!--#include file=Left.asp-->
	<%
	response.write  "</td>"&_
	  			    "		<td width=6> </td>"&_
					"		<td valign=top><table><tr><td height=20 valign=top><a href=index.asp>首页</a>&nbsp; &raquo; &nbsp;"&nav&" </td></tr></table>"&_
					"		<table width='100%' cellspacing=0 cellpadding=4 style='border-collapse: collapse' class=righttable><tbody class=table_td>"&_
					"			<tr>"&_
				    "				<td colspan=3 class=RightHead>"&nowplace&"</td>"&_
				    "			</tr>"
end sub

'***********************************************
' 目的： 公共下部分
'***********************************************
sub down()
	response.write "		</tbody></table>"&_
	"		</td>"&_
	"	</tr>"&_
	"</table>"
	%>
	<!--#include file=End.asp-->		
	<%
	response.write "</body>"&_
	"</html>"
end sub
'***********************************************

'*****************************************************************************
' 目的：    按商品特性显示最新商品列表
' 输入：    flag,RowNums,Row
' 说明：    flag:商品特性(1=新 2=荐 3=特);  NumsPerRow:每行商品数;  Rows:行数;
'*****************************************************************************
Sub ProductIndexList(flag,NumsPerRow,Rows)
    if IsNumeric(NumsPerRow)=false or IsNumeric(Rows)=false then
        check="false"
    end if
    dim topnums
    topnums=NumsPerRow*Rows
    if check<>"false" then
        response.write "<table cellspacing=0 cellpadding=2 style='border-collapse: collapse' width='100%'>"
        response.write "<tr align=center>"
        
        select case flag
            case 1
            	title_txt="全部商品"
                sql="select top "&topnums&" id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,1) order by id desc"
            case 2
                title_txt="精品推荐"
                sql="select top "&topnums&" id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,2) order by id desc"
            case 3
                title_txt="特价商品"
                sql="select top "&topnums&" id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,3) order by id desc"
        end select

        response.write  "<table cellspacing=0 cellpadding=0 style='border-collapse: collapse' width='100%'>"&_
						"<tbody class=table_td>"&_
        				"<tr>"&_
						"<td class=RightHead><a href=Product_ListFlag.asp?flag="&flag&" class=U>最新"&title_txt&"</a></td>"&_
						"<td class=RightHead align=right><a href=Product_ListFlag.asp?flag="&flag&" class=U><span style='font-weight: 400;padding-right:8px;'>更多"&title_txt&"</span></a></td>"&_
						"</tr><tr align=center><td colspan=2><table width='100%' cellspacing=1 cellpadding=4 class=MainTable><tbody class=table_td><tr>"
        
        set rs=server.createobject("adodb.recordset")
        rs.open sql,conn,1,1
        if rs.eof then 
            response.write "<td align=center>对不起，暂时没有相关商品信息!</td></tr></table>"
        else
            i=1
            set id                  =rs(0)
            set Product_info_Name   =rs(1)
            set Product_info_PriceM =rs(2)
            set Product_info_PriceS =rs(3)
            set Product_info_PicS   =rs(4)
            xxx=1/NumsPerRow*100

            while not rs.eof
                
                response.write "<td align=center width="&xxx&"% >"&_
                			   "  <table width='100%' border=0 align=center cellpadding=0 cellspacing=0 style='border-collapse: collapse'>"&_
                			   "    <tr>"&_
                			   "      <td align=center valign=top>"&_
                 			   "        <table border=1 cellspacing=0 cellpadding=4 style='border-collapse: collapse' bordercolor='#E4E4E4'>"&_
                   			   "          <tr><td align=center><a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&Product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a></td></tr>"&_
                 			   "        </table>"&_
                 			   "        <a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a><br />"&_
                 			   "        市场价：￥"&formatnumber(Product_info_PriceM,2,-1)&"<br>"&_
                 			   "        网站价：<b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b>"&_
                 			   "      </td>"&_
                 			   "    </tr>"&_
                 			   "  </table>"&_
                 			   "</td>"

                if i mod NumsPerRow = 0 then
                    response.write "</tr>"
                end if
                
                 rs.movenext
                 i=i+1
            wend
            response.write  "</tbody></table></td></tr></table>"&_  
						    "<div class=brclass></div>"  
        end if
    else
        response.write "参数错误"
    end if
    rs.close
    set rs=nothing
end sub
'*********************************************************



'*************************************************************************************************
' 目的：    按商品类别-显示商品列表
' 输入：    Bid,Sid,NumsPerPage,NumsPerRown,SortBy,showlist,flag
' 说明：    Bid:商品大类别id;  Sid:商品小类别id;  NumsPerPage:每页记录条数;  NumsPerRow:每行显示的商品数量; SortBy表示信息排序;  showlist表示商品显示方式;flag表示商品特性

'*************************************************************************************************
Sub Product_ListCategory(Bid,Sid,NumsPerRow,NumsPerPage)
    if IsNumeric(Bid)=false or IsNumeric(Sid)=false or IsNumeric(NumsPerRow)=false or IsNumeric(NumsPerPage)=false or IsNumeric(showlist)=false then
        check="false"
    end if
    if check<>"false" then
        response.write "<table cellspacing=0 cellpadding=2 style='border-collapse: collapse' width='100%'>"
        response.write "<tr align=center>"
        if flag=0 then
           if sid<>0 then
       	   	   sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and sid="&Sid&"  "&SortBy&" "
           else
         	   sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and bid="&Bid&"  "&SortBy&" "
           end if
        else
           if sid<>0 then
       	   	   sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,"&flag&") and sid="&Sid&"  "&SortBy&" "
           else
         	   sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,"&flag&") and bid="&Bid&"  "&SortBy&" "
           end if
        end if   
        
        set rs=server.createobject("adodb.recordset")
        rs.open sql,conn,1,1
        if rs.eof then 
            response.write "<td align=center>对不起，暂时没有相关商品信息!</td></tr></table>"
        else
            rs.PageSize =NumsPerPage '每页记录条数
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
            
            if showlist=1 then
                response.write  "<table cellspacing=1 cellpadding=5 width='100%' align=center>"&_
                				"<tr align=center>"
            else
                response.write "<table cellspacing=1 cellpadding=5 width='100%' align=center>"
            end if     
                        
            set id                  =rs(0)
            set Product_info_Name   =rs(1)
            set Product_info_PriceM =rs(2)
            set Product_info_PriceS =rs(3)
            set Product_info_PicS   =rs(4)
            xxx=1/NumsPerRow*100
            
            if showlist=2 then
                response.write "<tr><td>商品图片</td><td>商品名称(点击进入查看详细信息)</td><td>市场价</td><td>网站价</td></tr>"
            end if
            if showlist=3 then
                response.write "<tr><td>商品名称(点击进入查看详细信息)</td><td>市场价</td><td>网站价</td></tr>"
            end if

            while not rs.eof and i<=rs.pagesize
            
            select case showlist
            case 1 '图片方式显示
                response.write "<td align=center width="&xxx&"% >"&_
                               "  <table width='100%' border=0 align=center cellpadding=0 cellspacing=0 style='border-collapse: collapse'>"&_
                               "    <tr>"&_
                               "      <td align=center valign=top>"&_
                               "        <table border=1 cellspacing=0 cellpadding=4 style='border-collapse: collapse' bordercolor='#E4E4E4'>"&_
                               "          <tr><td align=center><a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&Product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a></td></tr>"&_
                               "        </table>"&_
                               "        <a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a><br />"&_
                               "        市场价：￥"&formatnumber(Product_info_PriceM,2,-1)&"<br>"&_
                               "        网站价：<b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b>"&_
                               "      </td>"&_
                               "    </tr>"&_
                               "  </table>"&_
                               "</td>"

                if i mod NumsPerRow = 0 then
                    response.write "</tr>"
                    response.write "<tr><td colspan="&NumsPerRow&" background=images/list_dotline.gif height=6></td></tr>"
                end if
                
            case 2  '列表方式显示
                response.write "<tr><td>"&_
               				   "        <a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a>"&_
                			   "    </td>"&_
		        			   "	<td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        			   "    <td>￥"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        			   "	<td><b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            			   "</tr>"&_
                			   "<tr><td colspan=4 background=images/list_dotline.gif height=6></td></tr>"
                
            case 3  '纯文字方式显示 
                response.write  "<tr><td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        				"<td>￥"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        				"<td><b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            				"</tr>"&_
	            				"<tr><td colspan=3 background=images/list_dotline.gif height=6></td></tr>"
            end select
                
            rs.movenext
            i=i+1
            wend
            call PageControl(iCount,maxpage,page)
            response.write "</table>"    
        end if
    else
        response.write "参数错误"
    end if
    rs.close
    set rs=nothing
end sub
'*********************************************************

'*************************************************************************************************
' 目的：    按商品特性-显示商品列表
' 输入：    flag,NumsPerPage,NumsPerRow
' 说明：    flag:商品特性(1=新 2=荐 3=特);  NumsPerPage:每页记录条数;  NumsPerRow:每行显示的商品数量; SortBy表示信息排序; showlist表示商品显示方式
'*************************************************************************************************
Sub Product_ListFlag(flag,NumsPerRow,NumsPerPage)
    if IsNumeric(NumsPerRow)=false or IsNumeric(NumsPerPage)=false or IsNumeric(showlist)=false then
        check="false"
    end if
    if check<>"false" then
        response.write "<table cellspacing=0 cellpadding=2 style='border-collapse: collapse' width='100%'>"
        response.write "<tr align=center>"
        
        select case flag
            case 1
                sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,1) "&SortBy&" "
            case 2
                sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,2) "&SortBy&" "
            case 3
                sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,3) "&SortBy&" "
        end select
        'response.write sql
        'response.end
        
        set rs=server.createobject("adodb.recordset")
        rs.open sql,conn,1,1
        if rs.eof then 
            response.write "<td align=center>对不起，暂时没有相关商品信息!</td></tr></table>"
        else
            rs.PageSize =NumsPerPage '每页记录条数
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

			if showlist=1 then
                response.write  "<table cellspacing=1 cellpadding=5 width='100%' align=center>"&_
                				"<tr align=center>"
            else
                response.write "<table cellspacing=1 cellpadding=5 width='100%' align=center>"
            end if     
                        
            set id                  =rs(0)
            set Product_info_Name   =rs(1)
            set Product_info_PriceM =rs(2)
            set Product_info_PriceS =rs(3)
            set Product_info_PicS   =rs(4)
            xxx=1/NumsPerRow*100
            
            if showlist=2 then
                response.write "<tr><td>商品图片</td><td>商品名称(点击进入查看详细信息)</td><td>市场价</td><td>网站价</td></tr>"
            end if
            if showlist=3 then
                response.write "<tr><td>商品名称(点击进入查看详细信息)</td><td>市场价</td><td>网站价</td></tr>"
            end if

            while not rs.eof and i<=rs.pagesize
            
            select case showlist
            case 1 '图片方式显示
                response.write "<td align=center width="&xxx&"% >"&_
                               "  <table width='100%' border=0 align=center cellpadding=0 cellspacing=0 style='border-collapse: collapse'>"&_
                               "    <tr>"&_
                               "      <td align=center valign=top>"&_
                               "        <table border=1 cellspacing=0 cellpadding=4 style='border-collapse: collapse' bordercolor='#E4E4E4'>"&_
                               "          <tr><td align=center><a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&Product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a></td></tr>"&_
                               "        </table>"&_
                               "        <a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a><br />"&_
                               "        市场价：￥"&formatnumber(Product_info_PriceM,2,-1)&"<br>"&_
                               "        网站价：<b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b>"&_
                               "      </td>"&_
                               "    </tr>"&_
                               "  </table>"&_
                               "</td>"

                if i mod NumsPerRow = 0 then
                    response.write "</tr>"
                    response.write "<tr><td colspan="&NumsPerRow&" background=images/list_dotline.gif height=6></td></tr>"
                end if
                
            case 2  '列表方式显示
                response.write "<tr><td>"&_
               				   "        <a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a>"&_
                			   "    </td>"&_
		        			   "	<td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        			   "    <td>￥"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        			   "	<td><b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            			   "</tr>"&_
                			   "<tr><td colspan=4 background=images/list_dotline.gif height=6></td></tr>"
                
            case 3  '纯文字方式显示 
                response.write  "<tr><td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        				"<td>￥"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        				"<td><b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            				"</tr>"&_
	            				"<tr><td colspan=3 background=images/list_dotline.gif height=6></td></tr>"
            end select                
            rs.movenext
            i=i+1
            wend
            call PageControl(iCount,maxpage,page)
            response.write "</table>"    
        end if
    else
        response.write "参数错误"
    end if
    rs.close
    set rs=nothing
end sub
'*********************************************************

'***********************************************
' 目的： 按商品搜索-显示商品列表
' 输入： Search,NumsPerPage,NumsPerRow
' 说明： Search:搜索参数集;  NumsPerPage:每页记录条数;  NumsPerRow:每行显示的商品数量; SortBy表示信息排序; showlist表示商品显示方式
'***********************************************
Sub Product_ListSearch(Search,NumsPerRow,NumsPerPage)
    if IsNumeric(NumsPerRow)=false or IsNumeric(NumsPerPage)=false or IsNumeric(showlist)=false then
        check="false"
    end if
    if check<>"false" then
        response.write "<table cellspacing=0 cellpadding=2 style='border-collapse: collapse' width='100%'>"
        response.write "<tr align=center>"
        
        if Search<>"" then
            sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 "&Search&SortBy
        else
            sql="select id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0"&SortBy
        end if
        
        set rs=server.createobject("adodb.recordset")
        rs.open sql,conn,1,1
        if rs.eof then 
            response.write "<td align=center>对不起，暂时没有相关商品信息!</td></tr></table>"
        else
            rs.PageSize =NumsPerPage '每页记录条数
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

			if showlist=1 then
                response.write  "<table cellspacing=1 cellpadding=5 width='100%' align=center>"&_
                				"<tr align=center>"
            else
                response.write "<table cellspacing=1 cellpadding=5 width='100%' align=center>"
            end if     
                        
            set id                  =rs(0)
            set Product_info_Name   =rs(1)
            set Product_info_PriceM =rs(2)
            set Product_info_PriceS =rs(3)
            set Product_info_PicS   =rs(4)
            xxx=1/NumsPerRow*100
            
            if showlist=2 then
                response.write "<tr><td>商品图片</td><td>商品名称(点击进入查看详细信息)</td><td>市场价</td><td>网站价</td></tr>"
            end if
            if showlist=3 then
                response.write "<tr><td>商品名称(点击进入查看详细信息)</td><td>市场价</td><td>网站价</td></tr>"
            end if

            while not rs.eof and i<=rs.pagesize
            
            select case showlist
            case 1 '图片方式显示
                response.write "<td align=center width="&xxx&"% >"&_
                               "  <table width='100%' border=0 align=center cellpadding=0 cellspacing=0 style='border-collapse: collapse'>"&_
                               "    <tr>"&_
                               "      <td align=center valign=top>"&_
                               "        <table border=1 cellspacing=0 cellpadding=4 style='border-collapse: collapse' bordercolor='#E4E4E4'>"&_
                               "          <tr><td align=center><a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&Product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a></td></tr>"&_
                               "        </table>"&_
                               "        <a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a><br />"&_
                               "        市场价：￥"&formatnumber(Product_info_PriceM,2,-1)&"<br>"&_
                               "        网站价：<b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b>"&_
                               "      </td>"&_
                               "    </tr>"&_
                               "  </table>"&_
                               "</td>"

                if i mod NumsPerRow = 0 then
                    response.write "</tr>"
                    response.write "<tr><td colspan="&RowNums&" background=images/list_dotline.gif height=6></td></tr>"
                end if
                
            case 2  '列表方式显示
                response.write "<tr><td>"&_
               				   "        <a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a>"&_
                			   "    </td>"&_
		        			   "	<td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        			   "    <td>￥"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        			   "	<td><b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            			   "</tr>"&_
                			   "<tr><td colspan=4 background=images/list_dotline.gif height=6></td></tr>"
                
            case 3  '纯文字方式显示 
                response.write  "<tr><td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        				"<td>￥"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        				"<td><b><font color=#FF6600>￥"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            				"</tr>"&_
	            				"<tr><td colspan=3 background=images/list_dotline.gif height=6></td></tr>"
            end select                
            rs.movenext
            i=i+1
            wend
            call PageControl(iCount,maxpage,page)
            response.write "</table>"    
        end if
    else
        response.write "参数错误"
    end if
    rs.close
    set rs=nothing
end sub
'***********************************************

'保存商品评论信息
sub Product_ReviewAddSave()
    dim prod_review_pid,prod_review_name,prod_review_detail,ErrMsg
    prod_review_pid   =my_request("prod_review_pid",1)
    prod_review_name  =my_request("prod_review_name",0)
    prod_review_detail=my_request("prod_review_detail",0)
    
    ErrMsg=""
    if prod_review_pid="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>评论商品ID号不能为空！</li>"
    end if
    if prod_review_name="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>评论人称呼不能为空！</li>"
    end if
    if prod_review_detail="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>评论/留言内容不能为空！</li>"
    end if

    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from prod_review"
        rs.open sql,conn,1,3
        rs.addnew
        rs("prod_review_pid")   =prod_review_pid
        rs("prod_review_name")  =prod_review_name
        rs("prod_review_detail")=prod_review_detail
        rs("prod_review_time")  =now()
        rs.update
        rs.close
        set rs=nothing
        call ok("恭喜，您已成功添加新评论！","Product_Detail.asp?id="&prod_review_pid&"")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub


'注册会员
sub User_RegSave()
    user_info_UserName =my_request("username",0)
    user_info_PassWord =my_request("PassWord",0)
    user_info_PassWord2=my_request("PassWord2",0)
    user_info_question =my_request("question",0)
    user_info_answer   =my_request("answer",0)
    user_info_RealName =my_request("realname",0)
    user_info_email    =my_request("email",0)
    user_info_sex      =my_request("sex",1)
    urlpath=my_request("urlpath",0)
    urlpath=replace(urlpath,"/","")
      
    ErrMsg=""
    if user_info_UserName="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>用户名不能为空！</li>"
    end if
    if user_info_PassWord="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>密码不能为空！</li>"
    end if
    if user_info_PassWord2="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>重复密码不能为空！</li>"
    end if
    if user_info_question="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>密保问题不能为空！</li>"
    end if
    if user_info_answer="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>问题答案不能为空！</li>"
    end if
        if  user_info_RealName="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>姓名不能为空！</li>"
    end if
    if user_info_email="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>电子邮件不能为空！</li>"
    end if
    if user_info_sex="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>性别不能为空！</li>"
    end if  

    if FoundErr<>True then
       	user_info_PassWord=md5(user_info_PassWord,32)
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from user_info"
        rs.open sql,conn,1,3
        rs.addnew
        rs("user_info_UserName")=user_info_UserName
        rs("user_info_PassWord")=user_info_PassWord
        rs("user_info_question")=user_info_question
        rs("user_info_answer")  =user_info_answer
        rs("user_info_RealName")=user_info_RealName
        rs("user_info_email")   =user_info_email
        rs("user_info_sex")     =user_info_sex
        rs("user_info_RegTime") =now()
        rs("user_info_LastLoginTime")=now()
        rs("user_info_LoginNums")=rs("user_info_LoginNums")+1
        rs.update
       
        session("user_info_id")=rs("user_info_id")
        session("user_info_UserName")=rs("user_info_UserName")
        session("user_info_LoginIn")=true
        rs.close
        set rs=nothing
        if urlpath<>"" then
            call ok("恭喜，您已成功注册成会员！",urlpath)
        else
            call ok("恭喜，您已成功注册成会员！","user_Personal.asp")
        end if
    else
	    call WriteErrMsg(ErrMsg)
    end if
end sub

'会员帐户资料修改
sub User_PersonalModiSave()
    user_info_RealName =my_request("user_info_RealName",0)
    user_info_email    =my_request("user_info_email",0)
    user_info_mobile   =my_request("user_info_mobile",0)
    user_info_address  =my_request("user_info_address",0)
    user_info_zip      =my_request("user_info_zip",0)
    
    ErrMsg=""
    if user_info_RealName="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>姓名不能为空！</li>"
    end if
    if user_info_address="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>收货地址不能为空！</li>"
    end if
    if user_info_mobile="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>联系电话不能为空！</li>"
    end if
        
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from user_info where user_info_id="&session("user_info_id")
        rs.open sql,conn,1,3
        rs("user_info_RealName")=user_info_RealName
        rs("user_info_email")   =user_info_email
        rs("user_info_mobile")=user_info_mobile
        rs("user_info_address")=user_info_address
        rs("user_info_zip")=user_info_zip
        rs.update
        rs.close
        set rs=nothing
        call ok("恭喜，您已成功更新会员个人资料！","user_Personal.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub

'会员密码-修改保存
sub User_PassWordModiSave() 
    passwordold=my_request("passwordold",0)
    password=my_request("password",0)
    confirmpassword=my_request("confirmpassword",0)
    
    ErrMsg=""
    if passwordold="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>旧密码不能为空！</li>"
    end if
    if password="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>新密码不能为空！</li>"
    end if
    if confirmpassword="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>重复新密码不能为空！</li>"
    end if        
    if password<>confirmpassword then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>新密码与重复新密码输入不一致！</li>"
    end if

    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from user_info where user_info_UserName='"&session("user_info_UserName")&"'"
        rs.open sql,conn,1,3
        password11=rs("user_info_PassWord")
        if password11<>md5(passwordold,32) then
            response.write"<SCRIPT language=JavaScript>alert('旧密码输入有错误！');"
            response.write"javascript:history.go(-1)</SCRIPT>"
            response.end
        else
            rs("user_info_PassWord")=md5(password,32)
            rs.update
        end if
        rs.close
        set rs=nothing
        Response.write "<script>alert(""您的密码已成功修改"");location.href=""user_PassWord.asp"";</script>"
        Response.end 
    else
        call WriteErrMsg(ErrMsg)
    end if      
end sub
%>