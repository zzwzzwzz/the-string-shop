<%
'***********************************************
' Ŀ�ģ� �����ϲ���
' ���룺 title,nowplace,nav
' ˵���� title:��ǰҳ����; nowplace:��ǰλ������; nav:����
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
					"		<td valign=top><table><tr><td height=20 valign=top><a href=index.asp>��ҳ</a>&nbsp; &raquo; &nbsp;"&nav&" </td></tr></table>"&_
					"		<table width='100%' cellspacing=0 cellpadding=4 style='border-collapse: collapse' class=righttable><tbody class=table_td>"&_
					"			<tr>"&_
				    "				<td colspan=3 class=RightHead>"&nowplace&"</td>"&_
				    "			</tr>"
end sub

'***********************************************
' Ŀ�ģ� �����²���
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
' Ŀ�ģ�    ����Ʒ������ʾ������Ʒ�б�
' ���룺    flag,RowNums,Row
' ˵����    flag:��Ʒ����(1=�� 2=�� 3=��);  NumsPerRow:ÿ����Ʒ��;  Rows:����;
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
            	title_txt="ȫ����Ʒ"
                sql="select top "&topnums&" id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,1) order by id desc"
            case 2
                title_txt="��Ʒ�Ƽ�"
                sql="select top "&topnums&" id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,2) order by id desc"
            case 3
                title_txt="�ؼ���Ʒ"
                sql="select top "&topnums&" id,Product_info_Name,Product_info_PriceM,Product_info_PriceS,Product_info_PicS from Product_info where Product_info_OnOff=0 and instr(Product_info_flag,3) order by id desc"
        end select

        response.write  "<table cellspacing=0 cellpadding=0 style='border-collapse: collapse' width='100%'>"&_
						"<tbody class=table_td>"&_
        				"<tr>"&_
						"<td class=RightHead><a href=Product_ListFlag.asp?flag="&flag&" class=U>����"&title_txt&"</a></td>"&_
						"<td class=RightHead align=right><a href=Product_ListFlag.asp?flag="&flag&" class=U><span style='font-weight: 400;padding-right:8px;'>����"&title_txt&"</span></a></td>"&_
						"</tr><tr align=center><td colspan=2><table width='100%' cellspacing=1 cellpadding=4 class=MainTable><tbody class=table_td><tr>"
        
        set rs=server.createobject("adodb.recordset")
        rs.open sql,conn,1,1
        if rs.eof then 
            response.write "<td align=center>�Բ�����ʱû�������Ʒ��Ϣ!</td></tr></table>"
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
                 			   "        �г��ۣ���"&formatnumber(Product_info_PriceM,2,-1)&"<br>"&_
                 			   "        ��վ�ۣ�<b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b>"&_
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
        response.write "��������"
    end if
    rs.close
    set rs=nothing
end sub
'*********************************************************



'*************************************************************************************************
' Ŀ�ģ�    ����Ʒ���-��ʾ��Ʒ�б�
' ���룺    Bid,Sid,NumsPerPage,NumsPerRown,SortBy,showlist,flag
' ˵����    Bid:��Ʒ�����id;  Sid:��ƷС���id;  NumsPerPage:ÿҳ��¼����;  NumsPerRow:ÿ����ʾ����Ʒ����; SortBy��ʾ��Ϣ����;  showlist��ʾ��Ʒ��ʾ��ʽ;flag��ʾ��Ʒ����

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
            response.write "<td align=center>�Բ�����ʱû�������Ʒ��Ϣ!</td></tr></table>"
        else
            rs.PageSize =NumsPerPage 'ÿҳ��¼����
		    iCount=rs.RecordCount '��¼����
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
                response.write "<tr><td>��ƷͼƬ</td><td>��Ʒ����(�������鿴��ϸ��Ϣ)</td><td>�г���</td><td>��վ��</td></tr>"
            end if
            if showlist=3 then
                response.write "<tr><td>��Ʒ����(�������鿴��ϸ��Ϣ)</td><td>�г���</td><td>��վ��</td></tr>"
            end if

            while not rs.eof and i<=rs.pagesize
            
            select case showlist
            case 1 'ͼƬ��ʽ��ʾ
                response.write "<td align=center width="&xxx&"% >"&_
                               "  <table width='100%' border=0 align=center cellpadding=0 cellspacing=0 style='border-collapse: collapse'>"&_
                               "    <tr>"&_
                               "      <td align=center valign=top>"&_
                               "        <table border=1 cellspacing=0 cellpadding=4 style='border-collapse: collapse' bordercolor='#E4E4E4'>"&_
                               "          <tr><td align=center><a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&Product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a></td></tr>"&_
                               "        </table>"&_
                               "        <a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a><br />"&_
                               "        �г��ۣ���"&formatnumber(Product_info_PriceM,2,-1)&"<br>"&_
                               "        ��վ�ۣ�<b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b>"&_
                               "      </td>"&_
                               "    </tr>"&_
                               "  </table>"&_
                               "</td>"

                if i mod NumsPerRow = 0 then
                    response.write "</tr>"
                    response.write "<tr><td colspan="&NumsPerRow&" background=images/list_dotline.gif height=6></td></tr>"
                end if
                
            case 2  '�б�ʽ��ʾ
                response.write "<tr><td>"&_
               				   "        <a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a>"&_
                			   "    </td>"&_
		        			   "	<td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        			   "    <td>��"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        			   "	<td><b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            			   "</tr>"&_
                			   "<tr><td colspan=4 background=images/list_dotline.gif height=6></td></tr>"
                
            case 3  '�����ַ�ʽ��ʾ 
                response.write  "<tr><td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        				"<td>��"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        				"<td><b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
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
        response.write "��������"
    end if
    rs.close
    set rs=nothing
end sub
'*********************************************************

'*************************************************************************************************
' Ŀ�ģ�    ����Ʒ����-��ʾ��Ʒ�б�
' ���룺    flag,NumsPerPage,NumsPerRow
' ˵����    flag:��Ʒ����(1=�� 2=�� 3=��);  NumsPerPage:ÿҳ��¼����;  NumsPerRow:ÿ����ʾ����Ʒ����; SortBy��ʾ��Ϣ����; showlist��ʾ��Ʒ��ʾ��ʽ
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
            response.write "<td align=center>�Բ�����ʱû�������Ʒ��Ϣ!</td></tr></table>"
        else
            rs.PageSize =NumsPerPage 'ÿҳ��¼����
		    iCount=rs.RecordCount '��¼����
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
                response.write "<tr><td>��ƷͼƬ</td><td>��Ʒ����(�������鿴��ϸ��Ϣ)</td><td>�г���</td><td>��վ��</td></tr>"
            end if
            if showlist=3 then
                response.write "<tr><td>��Ʒ����(�������鿴��ϸ��Ϣ)</td><td>�г���</td><td>��վ��</td></tr>"
            end if

            while not rs.eof and i<=rs.pagesize
            
            select case showlist
            case 1 'ͼƬ��ʽ��ʾ
                response.write "<td align=center width="&xxx&"% >"&_
                               "  <table width='100%' border=0 align=center cellpadding=0 cellspacing=0 style='border-collapse: collapse'>"&_
                               "    <tr>"&_
                               "      <td align=center valign=top>"&_
                               "        <table border=1 cellspacing=0 cellpadding=4 style='border-collapse: collapse' bordercolor='#E4E4E4'>"&_
                               "          <tr><td align=center><a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&Product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a></td></tr>"&_
                               "        </table>"&_
                               "        <a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a><br />"&_
                               "        �г��ۣ���"&formatnumber(Product_info_PriceM,2,-1)&"<br>"&_
                               "        ��վ�ۣ�<b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b>"&_
                               "      </td>"&_
                               "    </tr>"&_
                               "  </table>"&_
                               "</td>"

                if i mod NumsPerRow = 0 then
                    response.write "</tr>"
                    response.write "<tr><td colspan="&NumsPerRow&" background=images/list_dotline.gif height=6></td></tr>"
                end if
                
            case 2  '�б�ʽ��ʾ
                response.write "<tr><td>"&_
               				   "        <a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a>"&_
                			   "    </td>"&_
		        			   "	<td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        			   "    <td>��"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        			   "	<td><b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            			   "</tr>"&_
                			   "<tr><td colspan=4 background=images/list_dotline.gif height=6></td></tr>"
                
            case 3  '�����ַ�ʽ��ʾ 
                response.write  "<tr><td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        				"<td>��"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        				"<td><b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
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
        response.write "��������"
    end if
    rs.close
    set rs=nothing
end sub
'*********************************************************

'***********************************************
' Ŀ�ģ� ����Ʒ����-��ʾ��Ʒ�б�
' ���룺 Search,NumsPerPage,NumsPerRow
' ˵���� Search:����������;  NumsPerPage:ÿҳ��¼����;  NumsPerRow:ÿ����ʾ����Ʒ����; SortBy��ʾ��Ϣ����; showlist��ʾ��Ʒ��ʾ��ʽ
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
            response.write "<td align=center>�Բ�����ʱû�������Ʒ��Ϣ!</td></tr></table>"
        else
            rs.PageSize =NumsPerPage 'ÿҳ��¼����
		    iCount=rs.RecordCount '��¼����
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
                response.write "<tr><td>��ƷͼƬ</td><td>��Ʒ����(�������鿴��ϸ��Ϣ)</td><td>�г���</td><td>��վ��</td></tr>"
            end if
            if showlist=3 then
                response.write "<tr><td>��Ʒ����(�������鿴��ϸ��Ϣ)</td><td>�г���</td><td>��վ��</td></tr>"
            end if

            while not rs.eof and i<=rs.pagesize
            
            select case showlist
            case 1 'ͼƬ��ʽ��ʾ
                response.write "<td align=center width="&xxx&"% >"&_
                               "  <table width='100%' border=0 align=center cellpadding=0 cellspacing=0 style='border-collapse: collapse'>"&_
                               "    <tr>"&_
                               "      <td align=center valign=top>"&_
                               "        <table border=1 cellspacing=0 cellpadding=4 style='border-collapse: collapse' bordercolor='#E4E4E4'>"&_
                               "          <tr><td align=center><a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&Product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a></td></tr>"&_
                               "        </table>"&_
                               "        <a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a><br />"&_
                               "        �г��ۣ���"&formatnumber(Product_info_PriceM,2,-1)&"<br>"&_
                               "        ��վ�ۣ�<b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b>"&_
                               "      </td>"&_
                               "    </tr>"&_
                               "  </table>"&_
                               "</td>"

                if i mod NumsPerRow = 0 then
                    response.write "</tr>"
                    response.write "<tr><td colspan="&RowNums&" background=images/list_dotline.gif height=6></td></tr>"
                end if
                
            case 2  '�б�ʽ��ʾ
                response.write "<tr><td>"&_
               				   "        <a href=Product_Detail.asp?id="&id&"><img border=0 src=UPloadpic/"&product_info_PicS&" onload='loaded(this,"&root_option_WidthSPic&","&root_option_HeighSPic&")' /></a>"&_
                			   "    </td>"&_
		        			   "	<td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        			   "    <td>��"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        			   "	<td><b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
	            			   "</tr>"&_
                			   "<tr><td colspan=4 background=images/list_dotline.gif height=6></td></tr>"
                
            case 3  '�����ַ�ʽ��ʾ 
                response.write  "<tr><td><b><a href=Product_Detail.asp?id="&id&">"&Product_info_Name&"</a></b></td>"&_
		        				"<td>��"&formatnumber(Product_info_PriceM,2,-1)&"</td>"&_
		        				"<td><b><font color=#FF6600>��"&FormatNumber(Product_info_PriceS,2,-1)&"</font></b></td>"&_
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
        response.write "��������"
    end if
    rs.close
    set rs=nothing
end sub
'***********************************************

'������Ʒ������Ϣ
sub Product_ReviewAddSave()
    dim prod_review_pid,prod_review_name,prod_review_detail,ErrMsg
    prod_review_pid   =my_request("prod_review_pid",1)
    prod_review_name  =my_request("prod_review_name",0)
    prod_review_detail=my_request("prod_review_detail",0)
    
    ErrMsg=""
    if prod_review_pid="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>������ƷID�Ų���Ϊ�գ�</li>"
    end if
    if prod_review_name="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>�����˳ƺ�����Ϊ�գ�</li>"
    end if
    if prod_review_detail="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>����/�������ݲ���Ϊ�գ�</li>"
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
        call ok("��ϲ�����ѳɹ���������ۣ�","Product_Detail.asp?id="&prod_review_pid&"")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub


'ע���Ա
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
	    ErrMsg=ErrMsg & "<li>�û�������Ϊ�գ�</li>"
    end if
    if user_info_PassWord="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>���벻��Ϊ�գ�</li>"
    end if
    if user_info_PassWord2="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>�ظ����벻��Ϊ�գ�</li>"
    end if
    if user_info_question="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>�ܱ����ⲻ��Ϊ�գ�</li>"
    end if
    if user_info_answer="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>����𰸲���Ϊ�գ�</li>"
    end if
        if  user_info_RealName="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>��������Ϊ�գ�</li>"
    end if
    if user_info_email="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>�����ʼ�����Ϊ�գ�</li>"
    end if
    if user_info_sex="" then
	    FoundErr=True
	    ErrMsg=ErrMsg & "<li>�Ա���Ϊ�գ�</li>"
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
            call ok("��ϲ�����ѳɹ�ע��ɻ�Ա��",urlpath)
        else
            call ok("��ϲ�����ѳɹ�ע��ɻ�Ա��","user_Personal.asp")
        end if
    else
	    call WriteErrMsg(ErrMsg)
    end if
end sub

'��Ա�ʻ������޸�
sub User_PersonalModiSave()
    user_info_RealName =my_request("user_info_RealName",0)
    user_info_email    =my_request("user_info_email",0)
    user_info_mobile   =my_request("user_info_mobile",0)
    user_info_address  =my_request("user_info_address",0)
    user_info_zip      =my_request("user_info_zip",0)
    
    ErrMsg=""
    if user_info_RealName="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��������Ϊ�գ�</li>"
    end if
    if user_info_address="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>�ջ���ַ����Ϊ�գ�</li>"
    end if
    if user_info_mobile="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>��ϵ�绰����Ϊ�գ�</li>"
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
        call ok("��ϲ�����ѳɹ����»�Ա�������ϣ�","user_Personal.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub

'��Ա����-�޸ı���
sub User_PassWordModiSave() 
    passwordold=my_request("passwordold",0)
    password=my_request("password",0)
    confirmpassword=my_request("confirmpassword",0)
    
    ErrMsg=""
    if passwordold="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>�����벻��Ϊ�գ�</li>"
    end if
    if password="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>�����벻��Ϊ�գ�</li>"
    end if
    if confirmpassword="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>�ظ������벻��Ϊ�գ�</li>"
    end if        
    if password<>confirmpassword then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>���������ظ����������벻һ�£�</li>"
    end if

    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from user_info where user_info_UserName='"&session("user_info_UserName")&"'"
        rs.open sql,conn,1,3
        password11=rs("user_info_PassWord")
        if password11<>md5(passwordold,32) then
            response.write"<SCRIPT language=JavaScript>alert('�����������д���');"
            response.write"javascript:history.go(-1)</SCRIPT>"
            response.end
        else
            rs("user_info_PassWord")=md5(password,32)
            rs.update
        end if
        rs.close
        set rs=nothing
        Response.write "<script>alert(""���������ѳɹ��޸�"");location.href=""user_PassWord.asp"";</script>"
        Response.end 
    else
        call WriteErrMsg(ErrMsg)
    end if      
end sub
%>