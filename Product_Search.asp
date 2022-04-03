<center><%
//最上部调出或写入
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp-->
<!--#include file=include/Pages.asp-->
<%
dim count
set rs=server.createobject("adodb.recordset")
sql = "select prod_smallclass_bid,prod_smallclass_id,prod_smallclass_name from prod_smallclass order by prod_smallclass_bid desc"
rs.open sql,conn,1,1
set prod_smallclass_bid =rs(0)
set prod_smallclass_id  =rs(1)
set prod_smallclass_name=rs(2)
%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
subcat[0] = new Array("此大类下所有小类","<%=prod_smallclass_bid%>","");
        <%
        count = 1
        do while not rs.eof 
        ss=prod_smallclass_bid
        %>
subcat[<%=count%>] = new Array("<%=prod_smallclass_name%>","<%=prod_smallclass_bid%>","<%=prod_smallclass_id%>");
        <%
        count = count + 1
        rs.movenext
        if prod_smallclass_bid<>ss then
        %>
subcat[<%=count%>] = new Array("此大类下所有小类","<%=prod_smallclass_bid%>","");   
        <%
        count = count + 1
        end if
        loop
        rs.close
        set rs=nothing
        %>
onecount=<%=count%>;

//类别切换
function changelocation(locationid)
    {
    document.form1.sid.length = 0; 

    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form1.sid.options[document.form1.sid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    } 
</script>
<%
call up("商品搜索","商品搜索","商品搜索")

//整体框架中部-主体
response.write "			<form name=form1 action=Product_ListSearch.asp method=get>"
response.write "			<tr><td valign=top>商品类别:</td>"
response.write "				<td>"
response.write "					<select size=5 name=bid onChange=changelocation(document.form1.bid.options[document.form1.bid.selectedIndex].value)>"
response.write "					<option value=''>请选择大类</option>"
		    						sql="select prod_bigclass_id,prod_bigclass_name from prod_bigclass order by prod_bigclass_id desc"
		    						set rs=conn.execute (sql)
		    						set prod_bigclass_id=rs(0)
		    						set prod_bigclass_name=rs(1)
		    						do while not rs.eof
		    						response.write "<option value="&prod_bigclass_id&">"&prod_bigclass_name&"</option>"
		    						rs.movenext
		    						loop
		    						rs.close
		    						set rs=nothing
response.write "					</select>"
response.write "					<select name=sid size=5>"
response.write "					<option value=''>请选择小类</option>"
response.write "					</select>"
response.write "		 		</td>"
response.write "			</tr>"
response.write "			<tr><td>商品名称包含:</td><td><input type=text name=name size=30> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
response.write "			<tr><td>商品内容包含:</td><td><input type=text name=detail size=30></td></tr>"
response.write "			<tr><td>商品价格(本站价):</td><td><input size=6 name=UserPriceMin>元(小值)&nbsp; 至 <input size=6 name=UserPriceMax>元(大值)</td></tr>"
response.write "			<tr><td>商品特性:</td><td><input type=checkbox name=flag1 value=1>推荐&nbsp; <input type=checkbox name=flag2 value=2>新品 <input type=checkbox name=flag value=3>特价</td></tr>"
response.write "			<tr><td>  </td><td><input class=button type=submit value='开始搜索'></td></tr>"
response.write "			</form>"

call down()
%></center>