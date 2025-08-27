<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim CNT1

dim arrSQL(1,1)

arrSQL(0,0) = "회원조회1 (V1 : ID)"
arrSQL(0,1) = "select * from tbMember where M_ID like '%$value1$%' order by M_ID asc"

arrSQL(1,0) = "매크로업로드조회 (V1 : YYYY-MM-DD, V2 : YYYY-MM-DD)"
arrSQL(1,1) = "select * from tbMacro_Log where ML_UploadDate between '$value1$' and '$value2$'"
%>

<div align="center">
<h2>Custom Query</h2>	
<Script language="javascript">
function searchFormSubmit()
{	
	frmList2Excel.submit();
}
</script>
<table border=1 width=600px>
<form name="frmList2Excel" action="cq_list2excel.asp" method="post" target="_blank">
<tr>
	<td colspan=2 align=center>
		<select name="SQL">
<%
for CNT1 = 0 to ubound(arrSQL)
%>
			<option value="<%=arrSQL(CNT1,1)%>"><%=arrSQL(CNT1,0)%></option>
<%
next
%>
		</select>
		</td>
</tr>
<tr>
	<td>
		V1 : <input type="text" name=value1 size=20><br>
		V2 : <input type="text" name=value2 size=20><br>
		V3 : <input type="text" name=value3 size=20>
	</td>
	<td>
		<input type="button" value="지우기" style="width:70px" onclick="value1.value='';value2.value='';value3.value='';"><br>
		<input type="button" value="엑셀" style="width:70px" onclick="searchFormSubmit('xls');">
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