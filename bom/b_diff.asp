<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
call BOMVersion_Guide()
%>
<div align="center">
<h2>BOM Diff</h2>	
<Script language="javascript">
function searchFormSubmit()
{
	var nCountPNO = 0;
	var strMessage="";
	for (var i=0; i < frmList2Excel.PNO.length; i++) 
		if(frmList2Excel.PNO[i].value != "")
			nCountPNO += 1;
	
	if (nCountPNO < 2)
	{
		alert("비교할 파트넘버를 2개이상 입력해주세요.");
		return false;
	}
	
	var bByGuide = true;
	
	for (var i=0; i < frmList2Excel.PNO.length; i++) 
		if(frmList2Excel.PNO[i].value != "")
			if (frmList2Excel.PNO[i].value.indexOf("-") == -1)
				bByGuide = false;


	if(!bByGuide)
	{
		alert("파트넘버는 팝업가이드를 이용하여 입력해주세요.");
		return false;
	}

	frmList2Excel.submit();
} 
</script>
<table border=1 width=600px>

<iframe name="ifrmXLSDown" src="about:blank" frameborder=1 width=0px height=0px></iframe>
<form name="frmList2Excel" action="b_diff2Excel.asp" method="post" target="ifrmXLSDown">

<tr>
	<td>
<%
dim CNT1		
for CNT1 = 1 to 10
%>
		PartNo <%if CNT1 < 10 then%>0<%end if%><%=CNT1%> : <input type="text" name=PNO size=25 onclick="javascript:show_BOMVersion_Guide(this);"><br>
<%
next
%>
	</td>
	<td>
		<input type="button" value="down XLS" style="width:70px" onclick="searchFormSubmit('xls');">
	</td>
</tr>
</table>
</form>

</div>
</body>
</html> 


<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->