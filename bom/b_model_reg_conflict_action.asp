<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->

<% 
dim CNT1
dim SQL

dim strParts_P_P_No
dim strConflict
dim B_Code
dim arrConflict
dim arrConflict2

strParts_P_P_No	= Request("strParts_P_P_No")
strConflict		= Request("strConflict")
B_Code			= Request("B_Code")


strParts_P_P_No = "-" & replace(strParts_P_P_No,", ","-") & "-"
arrConflict = split(strConflict,"/||/")

'response.write strConflict & "<br>"

for CNT1 = 0 to ubound(arrConflict) - 1
	arrConflict2 = split(arrConflict(CNT1),"/|/")
	
	if instr(strParts_P_P_No,"-"&arrConflict2(0)&"-") > 0 then
		SQL = "update tbBOM_Qty set BQ_P_Desc='"&arrConflict2(2)&"', BQ_P_Spec='"&arrConflict2(3)&"', BQ_P_Maker='"&arrConflict2(4)&"' where BOM_B_Code = "&B_Code&" and Parts_P_P_No='"&arrConflict2(1)&"'"
		'response.write SQL & "<br>"
		sys_DBCon.execute(SQL)
	end if
next
%>


<form name="frmRedirect" action="db_load_action.asp">
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="Diff_YN" value="<%=Request("Diff_YN")%>">

</form>


<script language="javascript">
frmRedirect.submit();
</script>

<!-- #include virtual = "/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->