<!-- #include Virtual = "/mseth/header/asp_header.asp" -->
<!-- #include Virtual = "/mseth/header/db_header.asp" -->
<!-- include Virtual = "/mseth/header/html_header.asp" -->
<!-- #include Virtual = "/mseth/header/layout_full_header.asp" -->
<!-- #include Virtual = "/mseth/header/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Line

s_Work_Date = Request("s_Work_Date")
if s_Work_Date = "" then
	s_Work_Date = date()
end if

s_Line = Request("s_Line")
if s_Line = "" then
	s_Line = "3"
end if
%>
<html>
<head>
	
</head>
<body topmargin=0 leftmargin=0 bgcolor=black>

<table width=100% height=1000px cellpadding=0 cellspacing=1 bgcolor="white" style="color:white;font-size:42px;text-align:center;font-weight:bold">
<col></col>
<col width=150px></col>
<col width=150px></col>
<col width=150px></col>
<col width=300px></col>
<col width=150px></col>
<col width=17px></col>

<tr height=85px bgcolor=skyblue style="color:navy">
	<td>PART-NO</td>
	<td>PLAN</td>
	<td>INPUT</td>
	<td>RMND</td>
	<td>TIME</td>
	<td>RATE</td>	
	<td>&nbsp;</td>
</tr>

<tr>
	<td colspan=7><iframe src="mtr_ifrm_process_state_record.asp?s_Work_Date=<%=s_Work_Date%>&s_Line=<%=s_Line%>" width=100% height=100% frameborder=0 scrolling=yes></iframe></td>
</tr>
</table>
</body>
</html>



<!-- include Virtual = "/mseth/header/html_tail.asp" -->
<!-- #include Virtual = "/mseth/header/db_tail.asp" -->


	