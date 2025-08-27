<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
dim CNT1

dim strConflict
dim B_Code
dim arrConflict
dim arrConflict2

strConflict	= Request("strConflict")
B_Code		= Request("B_Code")

arrConflict = split(strConflict,"/||/")
%>
<br>
<br>
<br>
<table border=0 cellspacing=0 cellpadding=0 width=950px style=table-layout:fixed bgcolor="#999999" align=center>
<tr bgcolor=white>
	<td>다음은 충돌된 항목들입니다.<br>갱신할 항목을 선택한 후 확인을 클릭하여 주십시오.</td>
</tr>
</table>
<br>

<table border=0 cellspacing=1 cellpadding=0 width=950px style=table-layout:fixed bgcolor="#999999" align=center>
<form name="frmRedirect" action="b_model_reg_conflict_action.asp" method="post">
<input type="hidden" name="strConflict" value="<%=strConflict%>">
<input type="hidden" name="B_Code" value="<%=B_Code%>">
<input type="hidden" name="Diff_YN" value="<%=Request("Diff_YN")%>">
<tr bgcolor=white>
	<td width=50px>&nbsp;</td>
	<td width=100px>P_No</td>
	<td width=400px>기존 내용</td>
	<td width=400px>갱신될 내용</td>
</tr>
<%
for CNT1 = 0 to ubound(arrConflict) - 1
	arrConflict2 = split(arrConflict(CNT1),"/|/")
%>
<tr bgcolor=white>
	<td rowspan=3><input type="checkbox" name="strParts_P_P_No" value="<%=arrConflict2(1)%>" checked></td>
	<td rowspan=3><%=arrConflict2(1)%></td>
	<td <%if arrConflict2(5) <> arrConflict2(2) then%>style="color:red;"<%end if%>><%=arrConflict2(5)%></td>
	<td <%if arrConflict2(5) <> arrConflict2(2) then%>style="color:red;"<%end if%>><%=arrConflict2(2)%></td>
</tr>
<tr bgcolor=white>
	<td <%if arrConflict2(6) <> arrConflict2(3) then%>style="color:red;"<%end if%>><%=arrConflict2(6)%></td>
	<td <%if arrConflict2(6) <> arrConflict2(3) then%>style="color:red;"<%end if%>><%=arrConflict2(3)%></td>
</tr>
<tr bgcolor=white>
	<td <%if arrConflict2(7) <> arrConflict2(4) then%>style="color:red;"<%end if%>><%=arrConflict2(7)%></td>
	<td <%if arrConflict2(7) <> arrConflict2(4) then%>style="color:red;"<%end if%>><%=arrConflict2(4)%></td>
</tr>
<%
next
%>
</table>
<br>

<script language="javascript">
function SelectParts_P_P_NoAll()
{
	var ChangeTo = "";
	if (frmRedirect.btnSelectAll.value == "전체선택")
	{
		ChangeTo = true;
		frmRedirect.btnSelectAll.value = "선택해제";
	}
	else
	{
		ChangeTo = false;
		frmRedirect.btnSelectAll.value = "전체선택";
	}
	
	if(frmRedirect.strParts_P_P_No.length)
	{
		for(var i=0; i < frmRedirect.strParts_P_P_No.length; i++)
			frmRedirect.strParts_P_P_No[i].checked = ChangeTo;
	}
	else
	{
		frmRedirect.strParts_P_P_No.checked = ChangeTo;
	}
}
</script>

<table border=0 cellspacing=0 cellpadding=0 width=950px style=table-layout:fixed bgcolor="#999999" align=center>
<tr bgcolor=white>
	<td><input type="button" name="btnSelectAll" value="선택해제" onclick="SelectParts_P_P_NoAll()" style="width:70px">&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="확인" style="width:70px"></td>
</tr>
</table>
</form>



<!-- #include virtual="/header/layout_tail.asp" -->
<!-- #include virtual="/header/html_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->
