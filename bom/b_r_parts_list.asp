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
dim RS3

dim CNT1

dim s_Parts_P_P_No
dim strAppend
dim arrAppend
dim arrAppend2
dim bAppendYN

dim strResultType

s_Parts_P_P_No = request("s_Parts_P_P_No")
strResultType = request("strResultType")
%>

<div align="center">
<h2>정상부품명으로 R부품 조회</h2>	
<Script language="javascript">
function searchFormSubmit(strResultType)
{
	if (!searchForm.s_Parts_P_P_No.value)
		alert('Parts_P_P_No value is blank!')
	else
	{
		idResult.style.display = "none";
		alert("It will take a few minutes.\nWait Please.");
		location.href='b_r_parts_list.asp?s_Parts_P_P_No='+searchForm.s_Parts_P_P_No.value+'&strResultType='+strResultType;		
	}
}
</script>
<table border=1>
<form name="searchForm" method="get" action="b_r_parts_list.asp">
<tr>
	<td>
		<input type="text" name="s_Parts_P_P_No" style="width:150px" value="<%=s_Parts_P_P_No%>">
	</td>
	<td>
		<input type="button" value="Search" style="width:70px" onclick="searchFormSubmit('web');"><br><br>
		<input type="button" value="Reset" style="width:70px" onclick="s_Parts_P_P_No.value=''"><br><br>
		<input type="button" value="Excel" style="width:70px" onclick="searchFormSubmit('xls');">
	</td>
</tr>
</table>
</form>
<%
if s_Parts_P_P_No <> "" then
	
	set RS1 = Server.CreateObject("ADODB.RecordSet")
	set RS2 = Server.CreateObject("ADODB.RecordSet")
	set RS3 = Server.CreateObject("ADODB.RecordSet")
	
	
	SQL = "select distinct B_D_No = case left(BOM_Sub_BS_D_No,3) "
	SQL = SQL & "		when '478' then  "
	SQL = SQL & "			left(BOM_Sub_BS_D_No,10) "
	SQL = SQL & "		when '499' then  "
	SQL = SQL & "			left(BOM_Sub_BS_D_No,10) "
	SQL = SQL & "		when '687' then  "
	SQL = SQL & "			left(BOM_Sub_BS_D_No,10) "
	SQL = SQL & "		else  "
	SQL = SQL & "			left(BOM_Sub_BS_D_No,9) "
	SQL = SQL & "		end "
	SQL = SQL & "from  "
	SQL = SQL & "	tbBOM_Qty "
	SQL = SQL & "where Parts_P_P_No = '"&s_Parts_P_P_No&"'"
	RS1.Open SQL,sys_DBCon
	
	do until RS1.Eof
		if len(RS1("B_D_No")) = 10 then
			SQL = "select top 1 BOM_Sub_BS_D_No from tbBOM_Qty where left(BOM_Sub_BS_D_No,10)='"&RS1("B_D_No")&"' "
		else
			SQL = "select top 1 BOM_Sub_BS_D_No from tbBOM_Qty where left(BOM_Sub_BS_D_No,9)='"&RS1("B_D_No")&"' "
		end if
		
		RS2.Open SQL,sys_DBCon
		
		SQL = "select * from tbBOM_Qty where BOM_Sub_BS_D_No = '"&RS2("BOM_Sub_BS_D_No")&"' order by BQ_Code"
		RS3.Open SQL,sys_DBCon
		bAppendYN = "N"
		do until RS3.Eof 
			if RS3("Parts_P_P_No") = s_Parts_P_P_No then
				bAppendYN = "Y"
			end if
			
			if RS3("BQ_Order") = "R" and bAppendYN = "Y" then
				if instr(strAppend,RS3("Parts_P_P_No")) = 0 then
					strAppend = strAppend & RS1("B_D_No") &"|/|"& RS3("Parts_P_P_No") & "|%|"
				end if
			end if
			
			if RS3("Parts_P_P_No") <> s_Parts_P_P_No and bAppendYN = "Y" and RS3("BQ_Order") <> "R" then
				bAppendYN = "N"
			end if
			
			RS3.MoveNext
		loop
		RS3.Close
		
		RS2.Close
		
		RS1.MoveNext
	loop
	RS1.Close
	
	set RS1 = nothing
	set RS2 = nothing
	set RS3 = nothing
end if
%>
<div id="idResult">
<table border width="600">
<tr align="center">
	<td bgcolor=skyblue width=200px>Assy PartNo</td>
	<td bgcolor=skyblue width=200px>Material PartNo</td>
	<td bgcolor=skyblue width=200px>Material R-PartNo</td>
</tr>
<%
arrAppend = split(strAppend,"|%|")
for CNT1 = 0 to ubound(arrAppend)-1
	arrAppend2 = split(arrAppend(CNT1),"|/|")
%>
<tr>
	<td><%=arrAppend2(0)%></td>
	<td><%=s_Parts_P_P_No%></td>
	<td><%=arrAppend2(1)%></td>
</tr>
<%
next
%>
</table>
</div>
<%
if strAppend <> "" and strResultType="xls" then
%>
<form name="frmList2Excel" action="b_r_parts_list2excel.asp" method="post" target="_blank" >
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