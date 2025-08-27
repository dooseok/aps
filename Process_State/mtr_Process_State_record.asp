<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Line
dim s_Jaje_YN

s_Work_Date = Request("s_Work_Date")
if s_Work_Date = "" then
	s_Work_Date = date()
end if

s_Line = Request("s_Line")
if s_Line = "" then
	s_Line = "PCBA1"
end if

s_Jaje_YN = Request("s_JaJe_YN")
if s_Jaje_YN = "" then
	s_Jaje_YN = "N"
end if
%>
<html>
<head>
	
</head>
<body topmargin=0 leftmargin=0 bgcolor=black>

<table width=100% height=1000px cellpadding=0 cellspacing=1 bgcolor="white" style="color:white;font-size:42px;text-align:center;font-weight:bold">
<%
if s_Jaje_YN = "Y" then
%>
<form name="frmLine" method="post" action="mtr_Process_State_Record.asp">
<input type="hidden" name="s_Jaje_YN" value="Y">
<tr height=35px bgcolor=skyblue style="color:navy">
	<td colspan=7 align=left>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<select name="s_Line" onchange="location.href='mtr_Process_State_Record.asp?s_line='+this.value+'&s_Jaje_YN=Y';" style="height:90%;font-size:33px">
		<option value="PCBA1"<%if s_Line="PCBA1" then%> selected<%end if%>>PCBA1</option>
		<option value="PCBA2"<%if s_Line="PCBA2" then%> selected<%end if%>>PCBA2</option>
		<option value="PCBA3"<%if s_Line="PCBA3" then%> selected<%end if%>>PCBA3</option>
		<option value="PCBA4"<%if s_Line="PCBA4" then%> selected<%end if%>>PCBA4</option>
		<option value="PCBA5"<%if s_Line="PCBA5" then%> selected<%end if%>>PCBA5</option>
		<option value="PCBA6"<%if s_Line="PCBA6" then%> selected<%end if%>>PCBA6</option>
		<option value="PCBA7"<%if s_Line="PCBA7" then%> selected<%end if%>>PCBA7</option>
		<option value="CBOX1"<%if s_Line="CBOX1" then%> selected<%end if%>>CBOX1</option>
		<option value="CBOX2"<%if s_Line="CBOX2" then%> selected<%end if%>>CBOX2</option>
		<option value="CBOX3"<%if s_Line="CBOX3" then%> selected<%end if%>>CBOX3</option>
		<option value="CBOX4"<%if s_Line="CBOX4" then%> selected<%end if%>>CBOX4</option>
		<option value="CBOX5"<%if s_Line="CBOX5" then%> selected<%end if%>>CBOX5</option>
		<option value="CBOX6"<%if s_Line="CBOX6" then%> selected<%end if%>>CBOX6</option>
		<option value="CBOX7"<%if s_Line="CBOX7" then%> selected<%end if%>>CBOX7</option>
		</select>
		<input type="submit" value="새로고침" style="height:90%">
	</td>
</tr>
</form>
<tr height=50px bgcolor=skyblue style="color:navy">
	<td width=351px>파트넘버</td>
	<td width=200px>계획</td>
	<td width=200px>실적</td>
	<td width=200px>잔량</td>
	<td>달성률</td>
</tr>
<%
else
%>
<tr height=85px bgcolor=skyblue style="color:navy">
	<td width=351px>파트넘버</td>
	<td width=200px>계획</td>
	<td width=200px>실적</td>
	<td width=200px>잔량</td>
	<td>달성률</td>
</tr>
<%
end if
%>
<tr>
	<td colspan=7><iframe src="mtr_ifrm_process_state_record.asp?s_Jaje_YN=<%=s_Jaje_YN%>&s_Work_Date=<%=s_Work_Date%>&s_Line=<%=s_Line%>&runtime_YN=<%=Request("runtime_YN")%>" width=100% height=100% frameborder=0 scrolling=yes></iframe></td>
</tr>
</table>
</body>
</html>



<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	