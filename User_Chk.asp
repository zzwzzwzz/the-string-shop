<%
'�жϴӺδ�ת����Ա��¼��֤ҳ��,�����������Ա��Ա��¼���Զ�ת�ؼ���ԭ����
if session("user_info_LoginIn")="" then

    select case nowplace
    case "add_fav"  '��Ʒ�ղ�ʱ,���л�Ա��¼��֤
        response.redirect "User_Login.asp?urlpath=Product_Favorite.asp?id="&id&""
        response.end
    case "add_order" '��Ʒ����ʱ,���л�Ա��¼��֤
        response.redirect "User_Login.asp?urlpath=Cart_Order.asp"
        response.end
    case else
        response.redirect "User_Login.asp"
        response.end      
    end select
    
end if
%>