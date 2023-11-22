<center><!--#include file="User_Chk.asp"-->
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<script language="JavaScript">   
//ȫѡ����    
function CheckAll(form) {
    for (var i=0;i<form.elements.length;i++)
    {
        var e = form.elements[i];
        if (e.name != 'chkall') e.checked = form.chkall.checked; 
    }
 }
</script>
<%
action=my_request("action",0)
if action="��ѡ������Ʒ������ղؼ�" then
    call del()
end if

'���̣�����ɾ��
sub del()
    prod_favorite_id=my_request("prod_favorite_id",0)
    if prod_favorite_id<>"" then
        pp=ubound(split(prod_favorite_id,","))+1 '�ж�����id�й��м�ά
        for v=1 to pp
            id=request("prod_favorite_id")(v)     
            conn.execute ("delete from [prod_favorite] where prod_favorite_id="&id)
        next
        call ok("��ѡ��Ϣ�ѳɹ�������ҵ��ղؼУ�","User_Fav.asp")
    end if
end sub

call up("�ҵ��ղؼ�","�ҵ��ղؼ�","�ҵ��ղؼ�")
%>
<!--#include file="User_Menu.asp"-->
<%
response.write  "<tr><td>"&_
				"<table border=1 width=100% cellpadding=4 cellspacing=1 style='border-collapse: collapse' bordercolor=#E4E4E4>"&_
				"	<tr><td><b>ѡ��</b></td><td><b>��Ʒ����(�����ϸ�鿴)</b></td><td><b>�г���</b></td><td><b>��վ��</b></td><td><b>���빺�ﳵ</b></td></tr>"&_
				"	<form name=form1 action=User_Fav.asp method=post>"
    				set rs=server.createobject("adodb.recordset")
    				sql="select prod_favorite_id,prod_favorite_pid from prod_favorite where prod_favorite_uid="&session("user_info_id")
    				rs.open sql,conn,1,1
    				if rs.eof then 
    				    response.write "<tr><td colspan=5 align=center>Ŀǰ�����ղ���Ʒ��Ϣ��</td></tr>"
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
      
        				set prod_favorite_id=rs(0)
        				set prod_favorite_pid=rs(1) 
        				while not rs.eof and i<=rs.pagesize
        
            				//������Ʒ����Ϣ
            				sql1="select product_info_name,product_info_PriceM,product_info_PriceS from product_info where id="&prod_favorite_pid
            				set rs1=conn.execute (sql1)
            				product_info_name   =rs1(0)
            				product_info_PriceM	=rs1(1)
            				product_info_PriceS =rs1(2)
            				
            				rs1.close
            				set rs1=nothing    
response.write  "	<tr>"&_
				"		<td><input type=checkbox name=prod_favorite_id value="&prod_favorite_id&"></td>"&_
				"		<td><a href=product_detail.asp?id="&prod_favorite_pid&" target=_blank>"&product_info_name&"</a></td>"&_
				"		<td>"&product_info_PriceM&"</td>"&_
				"		<td>"&product_info_PriceS&"</td>"&_
				"		<td><a href=Cart_Add.asp?id="&prod_favorite_pid&">���빺�ﳵ</a></td>"&_
				"	</tr>"
         				rs.movenext
         				i=i+1
         				wend
response.write  "	<tr>"&_
				"		<td colspan=5>"&_
				"		<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ "&_
				"		<input class=button type=submit name=action value=ɾ�� onclick={if(confirm('��ȷ��Ҫ���ղؼ������ѡ������Ϣ��')){this.document.form1.submit();return true;}return false;}></td>"&_
				"	</tr>"&_
				"	<tr>"&_
				"		<td colspan=5>"&_
				"		<font color=#C0C0C0>ע���ղؼ��ڵ���Ʒ����ʱ��Ϊһ���£�����ϵͳ���Զ����!</font></td>"&_
				"	</tr>"&_
				"	<input type=hidden name=pagenow value="&page&">"
    				    call PageControl(iCount,maxpage,page)
    				end if
    			    rs.close
    			    set rs=nothing
response.write  "	</form>"&_
				"</table>"&_
				"</td></tr>"
call down()
%></center>