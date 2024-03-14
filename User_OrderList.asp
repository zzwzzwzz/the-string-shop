<center><!--#include file="User_Chk.asp"-->
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
'ȡ������
id=session("user_info_id")

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select user_info_RealName,user_info_email,user_info_mobile,user_info_address,user_info_zip from user_info where user_info_id="&id
rs.open sql,conn,1,1
user_info_RealName=rs(0)
user_info_email=rs(1)
user_info_mobile=rs(2)
user_info_address=rs(3)
user_info_zip=rs(4)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call User_PersonalModiSave()
end if

call up("�ҵĶ���","�ҵĶ���","�ҵĶ���")
%>
<!--#include file="User_Menu.asp"-->
<%
response.write  "<tr><td>"&_

				"<table border=1 width=100% cellpadding=4 cellspacing=1 style='border-collapse: collapse' bordercolor=#E4E4E4>"&_
				"	<tr><td><b>�������</b></td><td><b>��������</b></td><td><b>�ֽ��ܶ�</b></td><td><b>����״̬</b></td><td><b>�鿴������ϸ</b></td></tr>"
    				set rs=server.createobject("adodb.recordset")
    				sql="select order_info_id,Order_info_no,order_info_BuyTime,order_info_AllCost,order_info_CheckStates from order_info where order_info_recycle=0 and order_info_uid="&session("user_info_id")&" order by order_info_id desc" 
    				rs.open sql,conn,1,1
    				if rs.eof then 
    				    response.write "<tr><td colspan=5 align=center><font color=red>Ŀǰ���޶�����Ϣ!</font></td></tr>"
    				else
      				    rs.PageSize =20 'ÿҳ��¼����
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
      
      			        set order_info_id=rs(0)
      			        set order_info_no=rs(1)
      			        set order_info_BuyTime=rs(2)
      			        set order_info_AllCost=rs(3)
      			        set order_info_CheckStates=rs(4)   

      			        while not rs.eof and i<=rs.pagesize
            
      			        select case order_info_CheckStates
          			        case 0
          			            order_info_CheckStatesTxt="�¶���(δȷ��)"
          			        case 1
          			            order_info_CheckStatesTxt="��Ա����ȡ��"
          			        case 2
          			            order_info_CheckStatesTxt="��Ч������ȡ��"
          			        case 3
          			            order_info_CheckStatesTxt="��ȷ�ϣ�������"
          			        case 4
           			            order_info_CheckStatesTxt="�ѷ��������ջ�"
          			        case 5
          			            order_info_CheckStatesTxt="����֧���ɹ�"
         			        case 6
           			           order_info_CheckStatesTxt="�������"
      			        end select
response.write  "	<tr>"&_
				"	    <td><a href=User_OrderDetail.asp?id="&order_info_id&">"&order_info_no&"</a></td>"&_
				"	    <td>"&order_info_BuyTime&"</td>"&_
				"	    <td>"&order_info_AllCost&"Ԫ</td>"&_
				"	    <td>"&order_info_CheckStatesTxt&"</td>"&_
				"	    <td><a href=User_OrderDetail.asp?id="&order_info_id&">�鿴������ϸ</a></td>"&_
				"	</tr>"
       			        rs.movenext
       			        i=i+1
       			        wend
       			        call PageControl(iCount,maxpage,page)
    			    end if
    			    rs.close
    			    set rs=nothing
response.write "</table>"

response.write "</td></tr>"
call down()
%></center>