<%
'ASP最新SQL防注入过滤函数
Function Checkstr(Str)  
    If Isnull(Str) Then  
        CheckStr = ""  
        Exit Function   
    End If  
    Str = Replace(Str,Chr(0),"", 1, -1, 1)  
    Str = Replace(Str, """", "&quot;", 1, -1, 1)  
    Str = Replace(Str,"<","&lt;", 1, -1, 1)  
    Str = Replace(Str,">","&gt;", 1, -1, 1)   
    Str = Replace(Str, "script", "&#115;cript", 1, -1, 0)  
    Str = Replace(Str, "SCRIPT", "&#083;CRIPT", 1, -1, 0)  
    Str = Replace(Str, "Script", "&#083;cript", 1, -1, 0)  
    Str = Replace(Str, "script", "&#083;cript", 1, -1, 1)  
    Str = Replace(Str, "object", "&#111;bject", 1, -1, 0)  
    Str = Replace(Str, "OBJECT", "&#079;BJECT", 1, -1, 0)  
    Str = Replace(Str, "Object", "&#079;bject", 1, -1, 0)  
    Str = Replace(Str, "object", "&#079;bject", 1, -1, 1)  
    Str = Replace(Str, "applet", "&#097;pplet", 1, -1, 0)  
    Str = Replace(Str, "APPLET", "&#065;PPLET", 1, -1, 0)  
    Str = Replace(Str, "Applet", "&#065;pplet", 1, -1, 0)  
    Str = Replace(Str, "applet", "&#065;pplet", 1, -1, 1)  
    Str = Replace(Str, "[", "&#091;")  
    Str = Replace(Str, "]", "&#093;")  
    Str = Replace(Str, """", "", 1, -1, 1)  
    Str = Replace(Str, "=", "&#061;", 1, -1, 1)  
    Str = Replace(Str, "’", "’’", 1, -1, 1)  
    Str = Replace(Str, "select", "sel&#101;ct", 1, -1, 1)  
    Str = Replace(Str, "execute", "&#101xecute", 1, -1, 1)  
    Str = Replace(Str, "exec", "&#101xec", 1, -1, 1)  
    Str = Replace(Str, "join", "jo&#105;n", 1, -1, 1)  
    Str = Replace(Str, "union", "un&#105;on", 1, -1, 1)  
    Str = Replace(Str, "where", "wh&#101;re", 1, -1, 1)  
    Str = Replace(Str, "insert", "ins&#101;rt", 1, -1, 1)  
    Str = Replace(Str, "delete", "del&#101;te", 1, -1, 1)  
    Str = Replace(Str, "update", "up&#100;ate", 1, -1, 1)  
    Str = Replace(Str, "like", "lik&#101;", 1, -1, 1)  
    Str = Replace(Str, "drop", "dro&#112;", 1, -1, 1)  
    Str = Replace(Str, "create", "cr&#101;ate", 1, -1, 1)  
    Str = Replace(Str, "rename", "ren&#097;me", 1, -1, 1)  
    Str = Replace(Str, "count", "co&#117;nt", 1, -1, 1)  
    Str = Replace(Str, "chr", "c&#104;r", 1, -1, 1)  
    Str = Replace(Str, "mid", "m&#105;d", 1, -1, 1)  
    Str = Replace(Str, "truncate", "trunc&#097;te", 1, -1, 1)  
    Str = Replace(Str, "nchar", "nch&#097;r", 1, -1, 1)  
    Str = Replace(Str, "char", "ch&#097;r", 1, -1, 1)  
    Str = Replace(Str, "alter", "alt&#101;r", 1, -1, 1)  
    Str = Replace(Str, "cast", "ca&#115;t", 1, -1, 1)  
    Str = Replace(Str, "exists", "e&#120;ists", 1, -1, 1)  
    Str = Replace(Str,Chr(13),"<br>", 1, -1, 1)  
    CheckStr = Replace(Str,"’","’’", 1, -1, 1)  
End Function

'过滤SQL非法字符
function checkStr1(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function

'高效的防SQL注入函数
FUNCTION CHECKSTR2(ISTR)
    DIM ISTR_FORM,SQL_KILL,SQL_KILL_1,SQL_KILL_2,ISTR_KILL 
    IF ISTR="" THEN EXIT FUNCTION
    ISTR=LCase(ISTR)
    ISTR_FORM=ISTR
    SQL_KILL="'|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare|set|;|from|="
    SQL_KILL_1=SPLIT(SQL_KILL,"|")
    
    FOR EACH SQL_KILL_2 IN SQL_KILL_1
        ISTR=REPLACE(ISTR,SQL_KILL_2,"")
    NEXT
    
    CHECKSTR=ISTR
    ISTR_KILL=REPLACE(ISTR_FORM,ISTR,"")
    IF ISTR<>ISTR_FORM THEN
        RESPONSE.WRITE "<script>alert('警告: 您提交的数据["&ISTR_FORM&"]中含有非法字符 ["&ISTR_KILL&"]');history.back();</Script>"
        RESPONSE.END
    END IF
END FUNCTION

'过滤表单字符
function HTMLcode(fString)
    if not isnull(fString) then
        fString = Replace(fString, CHR(13), "")
        fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
        fString = Replace(fString, CHR(10), "<BR>")
        HTMLcode = fString
    end if
end function

'过滤HTML代码
function HTMLEncode(fString)
    if not isnull(fString) then
        fString = replace(fString, ">", "&gt;")
        fString = replace(fString, "<", "&lt;")

        fString = Replace(fString, CHR(32), "&nbsp;")
        fString = Replace(fString, CHR(9), "&nbsp;")
        fString = Replace(fString, CHR(34), "&quot;")
        fString = Replace(fString, CHR(39), "&#39;")
        fString = Replace(fString, CHR(13), "")
        fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
        fString = Replace(fString, CHR(10), "<BR> ")

        HTMLEncode = fString
    end if
end function

Function leach(str) '去除部分html代码
    if str<>"" then
	    str=replace(replace(replace(replace(replace(str,chr(34),"&quot;"),chr(39),"&#039;"),"<","&lt;"),">","&gt;"),vbCrlf,"<br>")
    end if
	leach=str
End function

Function Outleach(str)
	if str<>"" then
	    str=replace(replace(replace(replace(replace(str,"&quot;",chr(34)),"&#039;",chr(39)),"&lt;","<"),"&gt;",">"),"<br>",vbCrlf)
	end if
    Outleach=str
End function
%>


