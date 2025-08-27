<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<% 
dim SQL
dim RS1
dim RS2

dim CNT1

dim s_Parts_P_P_No
dim arrParts_P_P_No
dim strAppend
dim arrAppend
dim arrAppend2
dim bAppendYN

dim strResultType

dim oldParts_P_P_No
dim nBQ_Qty

dim strReplacedParts_P_P_No

s_Parts_P_P_No = request("s_Parts_P_P_No")
strResultType = request("strResultType")

arrParts_P_P_No = split(s_Parts_P_P_No,chr(13)&chr(10))

strReplacedParts_P_P_No = replace(s_Parts_P_P_No,chr(13)&chr(10),"|")
%>

<div align="center">
<h2>R부품명으로 정상부품 조회</h2>	
<Script language="javascript">
function searchFormSubmit(strResultType)
{
	if (!searchForm.s_Parts_P_P_No.value)
		alert('Parts_P_P_No value is blank!')
	else
	{
		idResult.style.display = "none";
		alert("It will take a few minutes.\nWait Please.");
		searchForm.strResultType.value = strResultType;
		searchForm.submit();
	}
}
</script>
<table border=1>
<form name="searchForm" method="post" action="b_r_multi_parts_list.asp">
	<input type="hidden" name="strResultType" value="web">
<tr>
	<td>
		<textarea name="s_Parts_P_P_No" style="width:300px;height=150px"><%=Request("s_Parts_P_P_No")%></textarea>
	</td>
	<td>
		<input type="button" value="Search" style="width:70px" onclick="searchFormSubmit('web');"><br><br>
		<input type="button" value="Reset" style="width:70px" onclick="s_Parts_P_P_No.value=''"><br><br>
		<input type="button" value="Excel" style="width:70px" onclick="searchFormSubmit('xls');">
	</td>
</tr>
</form>
</table>

<%
if s_Parts_P_P_No <> "" then
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
	'for CNT1 = 0 to ubound(arrParts_P_P_No)	
		SQL = "select * from tbBOM_Qty order by BQ_Code desc"
		RS1.Open SQL,sys_DBCon
		bAppendYN = "N"
		do until RS1.Eof 
			if instr(strReplacedParts_P_P_No,RS1("Parts_P_P_No")) > 0 and RS1("BQ_Order")="R" then
				bAppendYN = "Y"
				oldParts_P_P_No = RS1("Parts_P_P_No")
			end if
			
			if RS1("BQ_Order") <> "R" and bAppendYN = "Y" then
				if RS1("BQ_Qty") > 0 then
					if instr(strAppend,RS1("BOM_Sub_BS_D_No")&"|/|"&RS1("Parts_P_P_No")&"|/|"&oldParts_P_P_No) = 0 then
						SQL = "select sum(BQ_Qty) from tbBOM_Qty where BOM_Sub_BS_D_No = '"&RS1("BOM_Sub_BS_D_No")&"' and Parts_P_P_No = '"&RS1("Parts_P_P_No")&"'"
						RS2.Open SQL,sys_DBCon
						nBQ_Qty = RS2(0)
						RS2.Close
						strAppend = strAppend & RS1("BOM_Sub_BS_D_No") &"|/|"& RS1("Parts_P_P_No") &"|/|"& oldParts_P_P_No &"|/|"& nBQ_Qty & "|%|"
					end if
				end if
				bAppendYN = "N"
				oldParts_P_P_No = ""
			end if
			
			RS1.MoveNext
		loop
		RS1.Close
	'next
	set RS1 = nothing
	set RS2 = nothing
end if
%>
<div id="idResult">
<table border width="700">
<tr align="center">
	<td bgcolor=pink width=200px>Assy PartNo</td>
	<td bgcolor=pink width=200px>Item PartNo</td>
	<td bgcolor=pink width=200px>Item R-PartNo</td>
	<td bgcolor=pink width=100px>Qty</td>
</tr>
<%
arrAppend = split(strAppend,"|%|")
for CNT1 = 0 to ubound(arrAppend)-1
	arrAppend2 = split(arrAppend(CNT1),"|/|")
%>
<tr>
	<td><%=arrAppend2(0)%></td>
	<td><%=arrAppend2(1)%></td>
	<td><%=arrAppend2(2)%></td>
	<td><%=arrAppend2(3)%></td>
</tr>
<%
next
%>
</table>
</div>

<%
if strAppend <> "" and strResultType="xls" then
%>
<form name="frmList2Excel" action="b_r_multi_parts_list2excel.asp" method="post" target="_blank" >
<input type="hidden" name="strAppend" value="<%=strAppend%>">
<input type="hidden" name="s_Parts_P_P_No" value="<%=s_Parts_P_P_No%>">
<input type="hidden" name="strFileName"		value="b_r_parts_list.asp">
</form>
<script language="javascript">
	frmList2Excel.submit();
</script>
<%
end if
%>

<script language="javascript">
	idResult.style.display = "block";
</script>


</div>
</body>
</html> 


<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->