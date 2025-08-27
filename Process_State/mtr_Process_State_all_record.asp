<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Line
dim s_Process

s_Process	= Request("s_Process")
s_Work_Date = Request("s_Work_Date")
if s_Work_Date = "" then
	s_Work_Date = date()
end if
%>
<html>
<head>

</head>
<body topmargin=0 leftmargin=0 bgcolor=black>

<table width=100% height=870px cellpadding=0 cellspacing=0 bgcolor="black" style="color:white;font-size:75px;text-align:center;font-weight:bold">
<tr>
	<td colspan=1><iframe src="mtr_ifrm_process_state_all_record.asp?s_Multi_YN=<%=Request("s_Multi_YN")%>&s_Process=<%=s_Process%>&s_Work_Date=<%=s_Work_Date%>" width=100% height=100% frameborder=0 scrolling=no></iframe></td>
</tr>
</table>
</body>
</html>





<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


	