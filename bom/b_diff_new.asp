<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
call BOMSubVersion_Guide()
%>
<div align="center">
<h2>BOM Diff</h2>	

<Script language="javascript">
function searchFormSubmit()
{	
	var strPNO1 = frmList2Excel.PNO1.value;
	var strPNO2 = frmList2Excel.PNO2.value;
	
	if(strPNO1 == "" || strPNO2 == "")
	{
		alert("파트넘버를 선택해주세요.");
		return false;
	}
	else if(count(strPNO1,"/") != 2 || count(strPNO2,"/") != 2)
	{
		alert("버젼 정보가 없습니다.\n팝업가이드를 이용하여 입력해주세요.");
		return false;
	}
	frmList2Excel.submit();
}

function count(main_str, sub_str) 
{
	main_str += '';
	sub_str += '';
	
	if (sub_str.length <= 0) 
	{
	    return main_str.length + 1;
	}

	subStr = sub_str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	return (main_str.match(new RegExp(subStr, 'gi')) || []).length;
}
</script>

<table border=1 width=600px>


<form name="frmList2Excel" action="b_diff_new2Excel.asp" method="post" target="ifrmXLSDown">

<tr>
	<td>
		PartNo 1 : <input type="text" name=PNO1 size=25 onclick="javascript:show_BOMSubVersion_Guide(this);"><br>
		PartNo 2 : <input type="text" name=PNO2 size=25 onclick="javascript:show_BOMSubVersion_Guide(this);">
	</td>
	<td>
		<input type="button" value="down XLS" style="width:70px" onclick="searchFormSubmit('xls');">
	</td>
</tr>
</table>
</form>
<iframe name="ifrmXLSDown" src="about:blank" frameborder=1 width=1000px height=700px></iframe>
</div>
</body>
</html>

<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->