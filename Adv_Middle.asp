<%
dim adv_middle_pic1,adv_middle_pic2,adv_middle_pic3,adv_middle_pic4,adv_middle_pic5
dim adv_middle_pic1Url,adv_middle_pic2Url,adv_middle_pic3Url,adv_middle_pic4Url,adv_middle_pic5Url
dim adv_middle_pic1Txt,adv_middle_pic2Txt,adv_middle_pic3Txt,adv_middle_pic4Txt,adv_middle_pic5Txt
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select adv_middle_pic1,adv_middle_pic2,adv_middle_pic3,adv_middle_pic4,adv_middle_pic5,adv_middle_pic1Url,adv_middle_pic2Url,adv_middle_pic3Url,adv_middle_pic4Url,adv_middle_pic5Url,adv_middle_pic1Txt,adv_middle_pic2Txt,adv_middle_pic3Txt,adv_middle_pic4Txt,adv_middle_pic5Txt from adv_middle where adv_middle_id=1"
rs.open sql,conn,1,1
adv_middle_pic1=rs(0)
adv_middle_pic2=rs(1)
adv_middle_pic3=rs(2)
adv_middle_pic4=rs(3)
adv_middle_pic5=rs(4)
adv_middle_pic1Url=rs(5)
adv_middle_pic2Url=rs(6)
adv_middle_pic3Url=rs(7)
adv_middle_pic4Url=rs(8)
adv_middle_pic5Url=rs(9)
adv_middle_pic1Txt=rs(10)
adv_middle_pic2Txt=rs(11)
adv_middle_pic3Txt=rs(12)
adv_middle_pic4Txt=rs(13)
adv_middle_pic5Txt=rs(14)
rs.close
set rs=nothing

if adv_middle_pic1<>"" then
    a="uploadpic/"&adv_middle_pic1
    b=""&adv_middle_pic1Url
    c=""&adv_middle_pic1Txt
end if
if adv_middle_pic2<>"" then
    a=a&"|uploadpic/"&adv_middle_pic2
    b=b&"|"&adv_middle_pic2Url
    c=c&"|"&adv_middle_pic2Txt
end if
if adv_middle_pic3<>"" then
    a=a&"|uploadpic/"&adv_middle_pic3
    b=b&"|"&adv_middle_pic3Url
    c=c&"|"&adv_middle_pic3Txt
end if
if adv_middle_pic4<>"" then
    a=a&"|uploadpic/"&adv_middle_pic4
    b=b&"|"&adv_middle_pic4Url
    c=c&"|"&adv_middle_pic4Txt
end if
if adv_middle_pic5<>"" then
    a=a&"|uploadpic/"&adv_middle_pic5
    b=b&"|"&adv_middle_pic5Url
    c=c&"|"&adv_middle_pic5Txt
end if

//首屏焦点图片开始
response.write  "<TABLE height=160 cellPadding=0 width=100% border=0 style=border-collapse: collapse class=table_td>"&_
				"	<TR>"&_
				"		<TD vAlign=top align=middle>"
%>
<script language = "JavaScript" type="text/javascript">
<!--     
var focus_width=462;
var focus_height=160;
var text_height=20;
var swf_height = focus_height+text_height;
      
var pics='<%=a%>'; //链接图片
var links='<%=b%>';//链接网址
var texts='<%=c%>';//链接文本说明
      
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="images/autoflash.swf"><param name=wmode value=transparent><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
document.write('<embed src="images/autoflash.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#ffffff" quality="high" width="'+ focus_width +'" height="'+ swf_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
document.write('</object>');     
//--> 
</SCRIPT>
<%
response.write  "		</TD>"&_
				"	</TR>"&_
				"</TABLE>"
%><!--首屏焦点图片结束-->




