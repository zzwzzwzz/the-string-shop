<%
'post����sqlע���������HTML������ʼ
function nosql(str)
    if not isnull(str) then
        str=trim(str)
        str=replace(str,";","&#59;")		'�ֺ�
        str=replace(str,"'","&#39;")		'������
        str=replace(str,"""","&quot;")		'˫����
        str=replace(str,"chr(9)","&nbsp;")	'�ո�
        str=replace(str,"chr(10)","<br>")	'�س�
        str=replace(str,"chr(13)","<br>")	'�س�
        str=replace(str,"chr(32)","&nbsp;")	'�ո�
        str=replace(str,"chr(34)","&quot;")	'˫����
        str=replace(str,"chr(39)","&#39;")	'������
        str=Replace(str, "script", "&#115cript")'jscript
        str=replace(str,"<","&lt;")	        '��<
        str=replace(str,">","&gt;")	        '��>
        str=replace(str,"(","&#40;")	        '��(
        str=replace(str,")","&#41;")	        '��)
        str=replace(str,"--","&#45;&#45;")	'SQLע�ͷ�
        nosql=str
    end if
end function
%>


