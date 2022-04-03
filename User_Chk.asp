<%
//判断从何处转到会员登录验证页的,记忆下来，以便会员登录后自动转回继续原操作
if session("user_info_LoginIn")="" then

    select case nowplace
    case "add_fav"  '商品收藏时,进行会员登录验证
        response.redirect "User_Login.asp?urlpath=Product_Favorite.asp?id="&id&""
        response.end
    case "add_order" '商品结帐时,进行会员登录验证
        response.redirect "User_Login.asp?urlpath=Cart_Order.asp"
        response.end
    case else
        response.redirect "User_Login.asp"
        response.end      
    end select
    
end if
%>