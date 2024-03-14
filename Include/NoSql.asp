<%
'post过滤sql注入代防范及HTML防护开始
function nosql(str)
    if not isnull(str) then
        str=trim(str)
        str=replace(str,";","&#59;")		'分号
        str=replace(str,"'","&#39;")		'单引号
        str=replace(str,"""","&quot;")		'双引号
        str=replace(str,"chr(9)","&nbsp;")	'空格
        str=replace(str,"chr(10)","<br>")	'回车
        str=replace(str,"chr(13)","<br>")	'回车
        str=replace(str,"chr(32)","&nbsp;")	'空格
        str=replace(str,"chr(34)","&quot;")	'双引号
        str=replace(str,"chr(39)","&#39;")	'单引号
        str=Replace(str, "script", "&#115cript")'jscript
        str=replace(str,"<","&lt;")	        '左<
        str=replace(str,">","&gt;")	        '右>
        str=replace(str,"(","&#40;")	        '左(
        str=replace(str,")","&#41;")	        '右)
        str=replace(str,"--","&#45;&#45;")	'SQL注释符
        nosql=str
    end if
end function
%>


