<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
order_no=my_request("search_order_info_no",0)
set rs=server.createobject("adodb.recordset")
sql="select order_info_CheckStates,order_info_CheckTime,order_info_BuyTime from order_info where order_info_no='"&order_no&"'"
rs.open sql,conn,1,1
if rs.eof then 
    CheckStates="û�����������,��ȷ��������Ķ������Ƿ�����!"
else
    order_info_CheckStates  =rs(0)
    order_info_CheckTime    =rs(1)
    order_info_BuyTime      =rs(2)
end if
rs.close
set rs=nothing

select case order_info_CheckStates
    case "0"
        CheckStates="�¶���(δȷ��)"
        order_info_CheckTime=order_info_BuyTime
    case "1"
        CheckStates="�˿�����ȡ��"
    case "2"
        CheckStates="��Ч������ȡ��"
    case "3"
        CheckStates="��ȷ�ϣ�������"
    case "4"
        CheckStates="�ѷ��������ջ�"
    case "5"
        CheckStates="����֧���ɹ�"
    case "6"
        CheckStates="�������"
end select

call up("������ѯ���","������ѯ���","������ѯ���")

response.write  "<tr><td>�����ţ� "&order_no&"</td></tr>"&_
				"<tr><td>����״̬�� <font color=red>" &CheckStates&"</font></td></tr>"&_
				"<tr><td>����ʱ�䣺 "&order_info_CheckTime&"</td></tr>"
call down()
%></center>