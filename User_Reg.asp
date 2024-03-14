<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file="include/md5.asp"-->
<!--#include file=Sub.asp -->
<%
urlpath=my_request("urlpath",0)

action=my_request("action",0)
if action="save" then
    call User_RegSave()
end if

call up("注册会员","注册会员","注册会员")
response.write  "<form name=form_reg action=user_reg.asp method=post>"&_
        		"<input type=hidden name=action value=save>"&_
        		"<input type=hidden name=urlpath value="&urlpath&">"&_

        		"<tr><td>&nbsp;用户名:</td><td><input type=text size=20 name=username>  <input class=button onclick=javascript:window.open('User_RegNameChk.asp?username='+form_reg.username.value,null,'width=60,height=40') href=# type=button value=Check></td></tr>"&_
        		"<tr><td>&nbsp;密  码:</td><td><input type=password size=20 name=password></td></tr>"&_
        		"<tr><td>&nbsp;重复密码:</td><td><input type=password size=20 name=password2></td></tr>"&_
        		"<tr><td>&nbsp;设置密保:</td>"&_
        		"<td>"&_
        		"<select name=question size=1>"&_
        		"		<option value='' selected>--请选择--</option>"&_
        		"		<option value=我的宠物名字？>我的宠物名字？</option>"&_
        		"		<option value=我最好的朋友是谁？>我最好的朋友是谁？</option>"&_
        		"		<option value=我最喜爱的颜色？>我最喜爱的颜色？</option>"&_
        		"		<option value=我最喜爱的电影？>我最喜爱的电影？</option>"&_
        		"		<option value=我最喜爱的影星？>我最喜爱的影星？</option>"&_
        		"		<option value=我最喜爱的歌曲？>我最喜爱的歌曲？</option>"&_
        		"		<option value=我最喜爱的食物？>我最喜爱的食物？</option>"&_
       		 	"		<option value=我最大的爱好？>我最大的爱好？</option>"&_
        		"		<option value=我中学校名全称是什么？>我中学校名全称是什么？</option>"&_
        		"		<option value=我的座右铭是？>我的座右铭是？</option>"&_
        		"		<option value=我最喜欢的小说的名字？>我最喜欢的小说的名字？</option>"&_
        		"		<option value=我最喜欢的卡通人物名字？>我最喜欢的卡通人物名字？</option>"&_
        		"		<option value=我母亲父亲的生日？>我母亲父亲的生日？</option>"&_
        		"		<option value=我最欣赏的一位名人的名字？>我最欣赏的一位名人的名字？</option>"&_
        		"		<option value=我最喜欢的运动队全称？>我最喜欢的运动队全称？</option>"&_
        		"		<option value=我最喜欢的一句影视台词？>我最喜欢的一句影视台词？</option>"&_
        		"</select>"&_
        		"</td></tr>"&_
        		"<tr><td>&nbsp;问题答案:</td><td><input type=text size=20 name=answer></td></tr>"&_
       		 	"<tr><td>&nbsp;姓名:</td><td><input type=text size=20 name=realname></td></tr>"&_
        		"<tr><td>&nbsp;性别:</td><td><input type=radio value=0 name=sex checked>先生&nbsp; &nbsp; <input type=radio value=1 name=sex>&nbsp; 女士</td></tr>"&_
        		"<tr><td>&nbsp;Email:</td><td><input type=text size=20 name=email></td></tr>"&_
        		"<tr><td></td><td><input class=button type=submit value='  提交  '></td></tr>"&_
        		"</form>"
call down()
%>
</center>